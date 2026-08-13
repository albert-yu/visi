// Exponential Triple Smoothing (ETS), the forecasting model behind Excel's
// FORECAST.ETS family.
//
// Excel documents this as the "AAA version" of ETS: Additive error,
// Additive trend, Additive seasonality -- i.e. Holt-Winters additive
// smoothing. The recurrences, with season length `m`:
//
//   level_t  = a*(y_t - s_{t-m}) + (1-a)*(level_{t-1} + trend_{t-1})
//   trend_t  = b*(level_t - level_{t-1}) + (1-b)*trend_{t-1}
//   season_t = g*(y_t - level_{t-1} - trend_{t-1}) + (1-g)*s_{t-m}
//
// and an h-step-ahead forecast of `level_n + h*trend_n + s_{n+h-m*ceil(h/m)}`.
// With `m <= 1` the seasonal terms drop out and this degrades to Holt's
// linear method, which is what Excel does when it detects no seasonality.

/// A timeline resolved onto a regular grid, plus the grid's own geometry.
pub struct Series {
    /// Observations, one per grid step, gaps already filled.
    pub values: Vec<f64>,
    /// First timeline value.
    pub start: f64,
    /// Constant spacing between consecutive grid points.
    pub step: f64,
}

/// A fitted model plus everything the STAT/CONFINT accessors need.
pub struct Model {
    pub alpha: f64,
    pub beta: f64,
    pub gamma: f64,
    pub period: usize,
    pub level: f64,
    pub trend: f64,
    /// Seasonal indices for the final period, oldest first.
    pub seasons: Vec<f64>,
    /// One-step-ahead in-sample residuals (actual - forecast).
    pub residuals: Vec<f64>,
    pub values: Vec<f64>,
    pub step: f64,
}

/// Excel's optimizer reports its smoothing parameters to three decimals and
/// never returns 0 or 1 for alpha/beta (a perfectly linear series comes back
/// as alpha = 0.9, beta = 0.001), so the search runs over the same
/// three-decimal grid within these bounds.
const PARAM_MIN: f64 = 0.001;
const PARAM_MAX: f64 = 0.9;
const PARAM_QUANTUM: f64 = 0.001;

fn quantize(x: f64) -> f64 {
    (x / PARAM_QUANTUM).round() * PARAM_QUANTUM
}

/// Collapses (timeline, values) pairs onto the regular grid ETS needs:
/// sorts by time, averages duplicate timestamps, infers the constant step,
/// and fills interior gaps.
///
/// `data_completion` 1 (Excel's default) interpolates a missing point as the
/// average of its neighbours; 0 treats it as a zero.
pub fn build_series(
    values: &[f64],
    timeline: &[f64],
    data_completion: bool,
) -> Result<Series, String> {
    if values.len() != timeline.len() {
        return Err("#N/A".to_string());
    }
    if values.len() < 2 {
        return Err("#VALUE!".to_string());
    }

    let mut pairs: Vec<(f64, f64)> = timeline
        .iter()
        .copied()
        .zip(values.iter().copied())
        .collect();
    pairs.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));

    // Average out duplicate timestamps (Excel's default aggregation).
    let mut times: Vec<f64> = Vec::new();
    let mut obs: Vec<f64> = Vec::new();
    for (t, v) in pairs {
        if let Some(last) = times.last().copied()
            && (t - last).abs() <= f64::EPSILON * last.abs().max(1.0)
        {
            let n = obs.len();
            // Running mean of the duplicates seen so far at this timestamp.
            obs[n - 1] = (obs[n - 1] + v) / 2.0;
            continue;
        }
        times.push(t);
        obs.push(v);
    }
    if times.len() < 2 {
        return Err("#VALUE!".to_string());
    }

    // The step is the smallest gap; every other gap must be a whole multiple
    // of it or the timeline isn't on a constant step at all.
    let mut step = f64::INFINITY;
    for w in times.windows(2) {
        let d = w[1] - w[0];
        if d <= 0.0 {
            return Err("#NUM!".to_string());
        }
        if d < step {
            step = d;
        }
    }
    if !step.is_finite() || step <= 0.0 {
        return Err("#NUM!".to_string());
    }

    let mut grid: Vec<Option<f64>> = Vec::new();
    let start = times[0];
    for (t, v) in times.iter().zip(obs.iter()) {
        let offset = (t - start) / step;
        let idx = offset.round();
        if (offset - idx).abs() > 1e-6 {
            return Err("#NUM!".to_string());
        }
        let idx = idx as usize;
        if idx >= grid.len() {
            grid.resize(idx + 1, None);
        }
        grid[idx] = Some(*v);
    }

    // Fill gaps. Excel tolerates up to 30% missing.
    let missing = grid.iter().filter(|g| g.is_none()).count();
    if missing * 10 > grid.len() * 3 {
        return Err("#NUM!".to_string());
    }
    let mut filled = Vec::with_capacity(grid.len());
    for i in 0..grid.len() {
        match grid[i] {
            Some(v) => filled.push(v),
            None => {
                if !data_completion {
                    filled.push(0.0);
                    continue;
                }
                let prev = filled.last().copied().unwrap_or(0.0);
                let next = grid[i + 1..]
                    .iter()
                    .flatten()
                    .next()
                    .copied()
                    .unwrap_or(prev);
                filled.push((prev + next) / 2.0);
            }
        }
    }

    Ok(Series {
        values: filled,
        start,
        step,
    })
}

/// Excel's automatic season-length detection. Returns 0 when the series
/// shows no repeating pattern.
///
/// Scores each candidate period by the autocorrelation of the
/// first-differenced series (differencing removes the trend, which would
/// otherwise swamp the seasonal signal and make every lag look correlated).
/// A candidate has to clear a correlation floor *and* beat every other
/// candidate to be reported.
pub fn detect_period(values: &[f64]) -> usize {
    let n = values.len();
    if n < 4 {
        return 0;
    }
    // First difference to detrend.
    let diffs: Vec<f64> = values.windows(2).map(|w| w[1] - w[0]).collect();
    let dn = diffs.len();
    let mean = diffs.iter().sum::<f64>() / dn as f64;
    let var: f64 = diffs.iter().map(|d| (d - mean) * (d - mean)).sum();
    if var <= 0.0 {
        return 0;
    }

    let max_period = (n / 2).min(8784);
    let mut best = (0usize, 0.0f64);
    for p in 2..=max_period {
        if dn <= p {
            break;
        }
        let mut acf = 0.0;
        for i in p..dn {
            acf += (diffs[i] - mean) * (diffs[i - p] - mean);
        }
        acf /= var;
        if acf > best.1 {
            best = (p, acf);
        }
    }
    // A genuine season repeats strongly; anything weaker is noise.
    if best.1 >= 0.3 { best.0 } else { 0 }
}

/// Least-squares line through `(i, ys[i])`, returned as `(intercept, slope)`.
fn linreg(ys: &[f64]) -> (f64, f64) {
    let n = ys.len() as f64;
    if n < 2.0 {
        return (ys.first().copied().unwrap_or(0.0), 0.0);
    }
    let mean_x = (n - 1.0) / 2.0;
    let mean_y = ys.iter().sum::<f64>() / n;
    let mut num = 0.0;
    let mut den = 0.0;
    for (i, &y) in ys.iter().enumerate() {
        let dx = i as f64 - mean_x;
        num += dx * (y - mean_y);
        den += dx * dx;
    }
    let slope = if den == 0.0 { 0.0 } else { num / den };
    (mean_y - slope * mean_x, slope)
}

/// Seeds level/trend/season, and reports how many leading observations were
/// consumed doing so.
///
/// Two details matter for a clean series to forecast exactly:
///  - Seasonal indices are measured against a fitted trend line, not against
///    each cycle's own mean. A per-cycle mean sits at the *centre* of its
///    cycle, so deseasonalizing with it leaves a piecewise-constant
///    staircase rather than a straight line, and the trend seeded from that
///    staircase is off by half a cycle of drift.
///  - The returned level is the state at index `warmup - 1`, i.e. just
///    before the first observation the recurrences will actually score.
///    Seeding at index 0 and *also* feeding observation 0 through the update
///    consumes that point twice, which alone is enough to stop a perfectly
///    linear series forecasting exactly.
fn initial_state(values: &[f64], period: usize) -> (f64, f64, Vec<f64>, usize) {
    let n = values.len();
    let m = period.max(1);

    let mut seasons = Vec::new();
    if period > 1 {
        // Classical decomposition: estimate the trend-cycle with a centred
        // moving average of exactly one season, then read the seasonal
        // indices off the residual.
        //
        // Fitting a straight line to the raw series instead does *not*
        // work: the seasonal pattern is not orthogonal to the linear basis,
        // so it biases the slope. On a series with true slope 0.5 and
        // indices [-1.75, 7.75, 2.25, -8.25] repeating, a plain
        // least-squares fit returns 0.353, and every seasonal index (and
        // hence the whole forecast) inherits that error. A centred moving
        // average of one full season averages the seasonality out by
        // construction.
        let half = m / 2;
        let mut sums = vec![0.0; m];
        let mut counts = vec![0.0; m];
        for i in half..n.saturating_sub(half) {
            // For an even season length the window straddles two points at
            // each end, which get half weight apiece.
            let trend_cycle = if m.is_multiple_of(2) {
                if i < half || i + half >= n {
                    continue;
                }
                let mut acc = 0.5 * values[i - half] + 0.5 * values[i + half];
                for v in &values[i - half + 1..i + half] {
                    acc += v;
                }
                acc / m as f64
            } else {
                let acc: f64 = values[i - half..=i + half].iter().sum();
                acc / m as f64
            };
            sums[i % m] += values[i] - trend_cycle;
            counts[i % m] += 1.0;
        }
        if counts.iter().all(|c| *c > 0.0) {
            seasons = sums.iter().zip(counts.iter()).map(|(s, c)| s / c).collect();
        } else {
            // Too short for a full moving average: fall back to deviations
            // from the overall mean, which is at least unbiased in level.
            let mean = values.iter().sum::<f64>() / n as f64;
            let mut acc = vec![0.0; m];
            let mut cnt = vec![0.0; m];
            for (i, &y) in values.iter().enumerate() {
                acc[i % m] += y - mean;
                cnt[i % m] += 1.0;
            }
            seasons = acc
                .iter()
                .zip(cnt.iter())
                .map(|(a, c)| if *c > 0.0 { a / c } else { 0.0 })
                .collect();
        }
        // Additive indices must sum to zero.
        let mean = seasons.iter().sum::<f64>() / m as f64;
        for season in seasons.iter_mut() {
            *season -= mean;
        }
    }

    let warmup = if period > 1 { m } else { 2 }
        .min(n.saturating_sub(1))
        .max(1);
    let deseasonalized: Vec<f64> = values[..warmup]
        .iter()
        .enumerate()
        .map(|(i, y)| y - seasons.get(i % m).copied().unwrap_or(0.0))
        .collect();
    let (intercept, slope) = linreg(&deseasonalized);
    let level = intercept + slope * (warmup as f64 - 1.0);
    (level, slope, seasons, warmup)
}

/// Runs the AAA recurrences for a fixed parameter triple, collecting the
/// one-step-ahead residuals the optimizer scores and STAT reports.
fn smooth(values: &[f64], period: usize, alpha: f64, beta: f64, gamma: f64) -> Model {
    let (mut level, mut trend, mut seasons, warmup) = initial_state(values, period);
    let m = period.max(1);
    if seasons.is_empty() {
        seasons = vec![0.0; m];
    }
    let mut residuals = Vec::with_capacity(values.len().saturating_sub(warmup));

    for (i, &y) in values.iter().enumerate().skip(warmup) {
        let s_idx = i % m;
        let season = seasons[s_idx];
        let forecast = level + trend + season;
        residuals.push(y - forecast);

        let prev_level = level;
        let deseasonalized = y - season;
        level = alpha * deseasonalized + (1.0 - alpha) * (level + trend);
        trend = beta * (level - prev_level) + (1.0 - beta) * trend;
        if period > 1 {
            seasons[s_idx] = gamma * (y - prev_level - trend) + (1.0 - gamma) * season;
        }
    }

    Model {
        alpha,
        beta,
        gamma,
        period,
        level,
        trend,
        seasons,
        residuals,
        values: values.to_vec(),
        step: 1.0,
    }
}

fn sse(values: &[f64], period: usize, alpha: f64, beta: f64, gamma: f64) -> f64 {
    // `residuals` already starts after the initialization window, so every
    // one of them is genuinely attributable to the parameters being scored.
    smooth(values, period, alpha, beta, gamma)
        .residuals
        .iter()
        .map(|r| r * r)
        .sum()
}

/// Fits alpha/beta/gamma by minimizing the in-sample one-step-ahead SSE.
///
/// Coordinate descent over progressively finer grids (0.1, then 0.01, then
/// 0.001) rather than one dense 3-D sweep -- the full three-decimal cube
/// inside the bounds would be ~7e8 evaluations, while this reaches the same
/// resolution in a few thousand.
pub fn fit(values: &[f64], period: usize) -> Model {
    let mut alpha = 0.5;
    let mut beta = 0.1;
    let mut gamma = if period > 1 { 0.1 } else { PARAM_MIN };

    for scale in [0.1, 0.01, 0.001] {
        for _ in 0..4 {
            let mut improved = false;
            for which in 0..3 {
                if which == 2 && period <= 1 {
                    continue;
                }
                let current = match which {
                    0 => alpha,
                    1 => beta,
                    _ => gamma,
                };
                // Scan in Excel's preference order so that a *tie* resolves
                // the way Excel's optimizer does. When the series fits
                // perfectly every parameter triple scores the same, and
                // Excel reports alpha at its maximum with beta at its
                // minimum (a perfectly linear input comes back as
                // alpha = 0.9, beta = 0.001), so alpha is scanned downward
                // from the top and beta/gamma upward from the bottom, with
                // only strict improvements accepted.
                let steps = ((PARAM_MAX - PARAM_MIN) / scale).round() as i64;
                let mut best = (current, f64::INFINITY);
                for k in 0..=steps {
                    let offset = k as f64 * scale;
                    let candidate = if which == 0 {
                        quantize(PARAM_MAX - offset)
                    } else {
                        quantize(PARAM_MIN + offset)
                    };
                    if !(PARAM_MIN - 1e-12..=PARAM_MAX + 1e-12).contains(&candidate) {
                        continue;
                    }
                    let score = match which {
                        0 => sse(values, period, candidate, beta, gamma),
                        1 => sse(values, period, alpha, candidate, gamma),
                        _ => sse(values, period, alpha, beta, candidate),
                    };
                    if score < best.1 - 1e-15 {
                        best = (candidate, score);
                    }
                }
                if (best.0 - current).abs() > 1e-12 {
                    improved = true;
                    match which {
                        0 => alpha = best.0,
                        1 => beta = best.0,
                        _ => gamma = best.0,
                    }
                }
            }
            if !improved {
                break;
            }
        }
    }

    if period <= 1 {
        // Excel reports gamma as an epsilon rather than a clean zero when
        // there is no seasonal component to smooth.
        gamma = f64::EPSILON;
    }
    let mut model = smooth(values, period, alpha, beta, gamma);
    model.gamma = gamma;
    model
}

impl Model {
    /// Forecast `h` steps past the end of the fitted series (h >= 1).
    pub fn forecast(&self, h: usize) -> f64 {
        let m = self.period.max(1);
        let n = self.values.len();
        let seasonal = if self.period > 1 {
            self.seasons[(n + h - 1) % m]
        } else {
            0.0
        };
        self.level + (h as f64) * self.trend + seasonal
    }

    /// Residual standard deviation, the basis for the prediction interval.
    fn residual_sd(&self) -> f64 {
        let tail = &self.residuals[..];
        if tail.len() < 2 {
            return 0.0;
        }
        let mean = tail.iter().sum::<f64>() / tail.len() as f64;
        let var =
            tail.iter().map(|r| (r - mean) * (r - mean)).sum::<f64>() / (tail.len() - 1) as f64;
        var.sqrt()
    }

    /// Half-width of the prediction interval `h` steps ahead.
    ///
    /// The interval widens with the horizon: for an additive-error model the
    /// h-step variance accumulates as `sigma^2 * (1 + (h-1)*(alpha^2 + ...))`,
    /// approximated here by the standard `1 + (h-1)*alpha^2` term.
    pub fn confint(&self, h: usize, confidence: f64) -> Result<f64, String> {
        if confidence <= 0.0 || confidence >= 1.0 {
            return Err("#NUM!".to_string());
        }
        let z = crate::core::stats::inv_normal_cdf(1.0 - (1.0 - confidence) / 2.0)?;
        let sd = self.residual_sd();
        let growth = 1.0 + (h.saturating_sub(1) as f64) * self.alpha * self.alpha;
        Ok(z * sd * growth.sqrt())
    }

    /// FORECAST.ETS.STAT's `statistic_type` values 1-8.
    pub fn stat(&self, which: usize) -> Result<f64, String> {
        let tail = &self.residuals[..];
        let actual = &self.values[self.values.len() - tail.len()..];
        let n = tail.len() as f64;
        match which {
            1 => Ok(self.alpha),
            2 => Ok(self.beta),
            3 => Ok(self.gamma),
            // MASE: mean absolute error scaled by the naive one-step error.
            4 => {
                if tail.is_empty() {
                    return Ok(0.0);
                }
                let mae = tail.iter().map(|r| r.abs()).sum::<f64>() / n;
                let naive: f64 = self.values.windows(2).map(|w| (w[1] - w[0]).abs()).sum();
                let denom = naive / (self.values.len() - 1).max(1) as f64;
                Ok(if denom == 0.0 { 0.0 } else { mae / denom })
            }
            // SMAPE: symmetric mean absolute percentage error.
            5 => {
                if tail.is_empty() {
                    return Ok(0.0);
                }
                let mut acc = 0.0;
                for (r, a) in tail.iter().zip(actual.iter()) {
                    let f = a - r;
                    let denom = (a.abs() + f.abs()) / 2.0;
                    if denom != 0.0 {
                        acc += (r.abs() / denom) / 2.0;
                    }
                }
                Ok(acc / n)
            }
            6 => Ok(if tail.is_empty() {
                0.0
            } else {
                tail.iter().map(|r| r.abs()).sum::<f64>() / n
            }),
            7 => Ok(if tail.is_empty() {
                0.0
            } else {
                (tail.iter().map(|r| r * r).sum::<f64>() / n).sqrt()
            }),
            8 => Ok(self.step),
            _ => Err("#NUM!".to_string()),
        }
    }
}

/// Shared front end for the whole FORECAST.ETS family: validates and
/// regularizes the timeline, resolves the season length, and fits.
///
/// `seasonality` follows Excel's convention -- 1 means "detect
/// automatically", 0 means "no seasonality", and anything else is an
/// explicit season length.
pub fn prepare(
    values: &[f64],
    timeline: &[f64],
    seasonality: f64,
    data_completion: bool,
) -> Result<Model, String> {
    if !(0.0..=8784.0).contains(&seasonality) {
        return Err("#NUM!".to_string());
    }
    let series = build_series(values, timeline, data_completion)?;
    let period = match seasonality.round() as i64 {
        1 => detect_period(&series.values),
        0 => 0,
        p => p as usize,
    };
    if period > series.values.len() {
        return Err("#NUM!".to_string());
    }
    let mut model = fit(&series.values, period);
    model.step = series.step;
    Ok(model)
}

/// Steps from the end of the fitted series to `target`, or an error when the
/// target is not strictly in the future.
pub fn horizon(series_start: f64, step: f64, n: usize, target: f64) -> Result<usize, String> {
    let last = series_start + step * (n as f64 - 1.0);
    if target <= last {
        return Err("#NUM!".to_string());
    }
    let h = ((target - last) / step).round();
    if h < 1.0 {
        return Err("#NUM!".to_string());
    }
    Ok(h as usize)
}
