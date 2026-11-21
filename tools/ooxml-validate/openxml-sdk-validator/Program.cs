using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var version = FileFormatVersions.Microsoft365;
var jsonOutput = false;
var files = new List<string>();

for (var i = 0; i < args.Length; i++)
{
    var arg = args[i];
    if (arg == "--json")
    {
        jsonOutput = true;
    }
    else if (arg == "--version")
    {
        if (i + 1 >= args.Length)
        {
            Console.Error.WriteLine("Missing value for --version");
            return 2;
        }
        i++;
        if (!TryParseVersion(args[i], out version))
        {
            Console.Error.WriteLine($"Unsupported --version value: {args[i]}");
            return 2;
        }
    }
    else if (arg.StartsWith("--version=", StringComparison.Ordinal))
    {
        var raw = arg["--version=".Length..];
        if (!TryParseVersion(raw, out version))
        {
            Console.Error.WriteLine($"Unsupported --version value: {raw}");
            return 2;
        }
    }
    else if (arg is "-h" or "--help")
    {
        Console.WriteLine("Usage: OpenXmlSdkValidator [--version <Office2007|Office2010|Office2013|Office2016|Office2019|Office2021|Microsoft365>] [--json] <file1.xlsx> [file2.xlsm ...]");
        return 0;
    }
    else if (arg.StartsWith("-", StringComparison.Ordinal))
    {
        Console.Error.WriteLine($"Unknown option: {arg}");
        return 2;
    }
    else
    {
        files.Add(arg);
    }
}

if (files.Count == 0)
{
    Console.Error.WriteLine("No input files provided.");
    return 2;
}

var validator = new OpenXmlValidator(version);
var results = new List<FileValidationResult>();
var anyFailed = false;

foreach (var filePath in files)
{
    var fileErrors = new List<ValidationErrorRecord>();
    if (!File.Exists(filePath))
    {
        fileErrors.Add(new ValidationErrorRecord(
            Id: "IO_NotFound",
            ErrorType: "Package",
            Part: "",
            XPath: "",
            Description: $"File not found: {filePath}"
        ));
    }
    else
    {
        try
        {
            using var doc = SpreadsheetDocument.Open(filePath, false);
            foreach (var err in validator.Validate(doc))
            {
                fileErrors.Add(new ValidationErrorRecord(
                    Id: err.Id ?? "",
                    ErrorType: err.ErrorType.ToString(),
                    Part: err.Part?.Uri?.ToString() ?? "",
                    XPath: err.Path?.XPath ?? "",
                    Description: err.Description ?? ""
                ));
            }
        }
        catch (Exception ex)
        {
            fileErrors.Add(new ValidationErrorRecord(
                Id: "OpenXmlException",
                ErrorType: "Package",
                Part: "",
                XPath: "",
                Description: $"{ex.GetType().Name}: {ex.Message}"
            ));
        }
    }

    if (fileErrors.Count > 0)
    {
        anyFailed = true;
    }

    results.Add(new FileValidationResult(filePath, fileErrors));
}

if (jsonOutput)
{
    var payload = new ValidationReport(version.ToString(), results);
    Console.WriteLine(JsonSerializer.Serialize(payload, new JsonSerializerOptions
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    }));
}
else
{
    foreach (var result in results)
    {
        if (result.Errors.Count == 0)
        {
            Console.WriteLine($"OK   {result.File}");
        }
        else
        {
            Console.WriteLine($"FAIL {result.File} ({result.Errors.Count} error(s)):");
            foreach (var err in result.Errors)
            {
                var loc = string.IsNullOrEmpty(err.Part) ? "" : $" [{err.Part}]";
                var xp = string.IsNullOrEmpty(err.XPath) ? "" : $" {err.XPath}";
                Console.WriteLine($"  - ({err.ErrorType}/{err.Id}){loc}{xp}: {err.Description}");
            }
        }
    }
}

return anyFailed ? 1 : 0;

static bool TryParseVersion(string raw, out FileFormatVersions parsed)
{
    if (!Enum.TryParse(raw, ignoreCase: true, out parsed))
    {
        return false;
    }
    return parsed is FileFormatVersions.Office2007
        or FileFormatVersions.Office2010
        or FileFormatVersions.Office2013
        or FileFormatVersions.Office2016
        or FileFormatVersions.Office2019
        or FileFormatVersions.Office2021
        or FileFormatVersions.Microsoft365;
}

sealed record ValidationErrorRecord(
    string Id,
    string ErrorType,
    string Part,
    string XPath,
    string Description
);

sealed record FileValidationResult(
    string File,
    List<ValidationErrorRecord> Errors
);

sealed record ValidationReport(
    string Version,
    List<FileValidationResult> Results
);
