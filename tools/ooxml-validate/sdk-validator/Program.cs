using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

static int MainEntry(string[] args)
{
    var jsonOutput = false;
    var version = FileFormatVersions.Microsoft365;
    var files = new List<string>();

    for (var i = 0; i < args.Length; i++)
    {
        switch (args[i])
        {
            case "--json":
                jsonOutput = true;
                break;
            case "--version":
                if (i + 1 >= args.Length)
                {
                    Console.Error.WriteLine("Missing value for --version");
                    return 2;
                }
                i++;
                if (!Enum.TryParse<FileFormatVersions>(args[i], ignoreCase: true, out version)
                    || version == FileFormatVersions.None)
                {
                    Console.Error.WriteLine($"Unknown FileFormatVersions value: {args[i]}");
                    return 2;
                }
                break;
            case "-h":
            case "--help":
                Console.WriteLine("Usage: OpenXmlSdkValidator [--json] [--version <FileFormatVersions>] <file.xlsx> [...]");
                return 0;
            default:
                if (args[i].StartsWith('-'))
                {
                    Console.Error.WriteLine($"Unknown option: {args[i]}");
                    return 2;
                }
                files.Add(args[i]);
                break;
        }
    }

    if (files.Count == 0)
    {
        Console.Error.WriteLine("Usage: OpenXmlSdkValidator [--json] [--version <FileFormatVersions>] <file.xlsx> [...]");
        return 2;
    }

    var validator = new OpenXmlValidator(version);
    var results = new List<FileValidationResult>(files.Count);
    var hasErrors = false;

    foreach (var file in files)
    {
        var errors = new List<ValidationIssue>();
        try
        {
            using var document = SpreadsheetDocument.Open(file, false);
            foreach (var error in validator.Validate(document))
            {
                errors.Add(new ValidationIssue(
                    Id: error.Id,
                    ErrorType: error.ErrorType.ToString(),
                    Description: error.Description,
                    Part: error.Part?.Uri?.ToString() ?? string.Empty,
                    Path: error.Path?.XPath ?? string.Empty,
                    Node: error.Node?.LocalName ?? string.Empty
                ));
            }
        }
        catch (Exception ex)
        {
            errors.Add(new ValidationIssue(
                Id: "OpenException",
                ErrorType: "Package",
                Description: $"{ex.GetType().Name}: {ex.Message}",
                Part: string.Empty,
                Path: string.Empty,
                Node: string.Empty
            ));
        }

        if (errors.Count > 0)
        {
            hasErrors = true;
        }
        results.Add(new FileValidationResult(file, errors));
    }

    if (jsonOutput)
    {
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true,
        };
        Console.WriteLine(JsonSerializer.Serialize(new ValidationReport(version.ToString(), results), options));
    }
    else
    {
        foreach (var result in results)
        {
            if (result.Errors.Count == 0)
            {
                Console.WriteLine($"OK: {result.File}");
            }
            else
            {
                Console.WriteLine($"FAIL: {result.File} ({result.Errors.Count} error(s))");
                foreach (var err in result.Errors)
                {
                    var location = string.IsNullOrEmpty(err.Part)
                        ? err.Path
                        : string.IsNullOrEmpty(err.Path)
                            ? err.Part
                            : $"{err.Part} {err.Path}";
                    Console.WriteLine($"  [{err.ErrorType}/{err.Id}] {location}: {err.Description}");
                }
            }
        }
    }

    return hasErrors ? 1 : 0;
}

return MainEntry(args);

sealed record ValidationIssue(
    string Id,
    string ErrorType,
    string Description,
    string Part,
    string Path,
    string Node
);

sealed record FileValidationResult(
    string File,
    List<ValidationIssue> Errors
);

sealed record ValidationReport(
    string Version,
    List<FileValidationResult> Results
);
