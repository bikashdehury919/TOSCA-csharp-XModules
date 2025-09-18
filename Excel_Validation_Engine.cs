using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml; // Make sure you installed EPPlus NuGet package
using Tricentis.Automation.AutomationInstructions.Configuration;
using Tricentis.Automation.AutomationInstructions.Dynamic.Values;
using Tricentis.Automation.AutomationInstructions.TestActions;
using Tricentis.Automation.Creation;
using Tricentis.Automation.Engines;
using Tricentis.Automation.Engines.SpecialExecutionTasks;
using Tricentis.Automation.Engines.SpecialExecutionTasks.Attributes;
using Tricentis.Automation.Execution.Results;

namespace ExcelValidationSET
{
    [SpecialExecutionTaskName("ValidateExcel")]
    public class ExcelValidationSET : SpecialExecutionTask
    {
        public ExcelValidationSET(Validator validator) : base(validator) { }

        public override ActionResult Execute(ISpecialExecutionTaskTestAction testAction)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                IInputValue filePathParam = testAction.GetParameterAsInputValue("ExcelPath", true);
                IInputValue sheetNameParam = testAction.GetParameterAsInputValue("SheetName", false);
                IInputValue mandatoryColsParam = testAction.GetParameterAsInputValue("MandatoryColumns", false);  // comma-separated
                IInputValue userParam = testAction.GetParameterAsInputValue("Name", false);
                IInputValue expectedEmailParam = testAction.GetParameterAsInputValue("ExpectedEmail", false);

                string filePath = filePathParam?.Value;
                string sheetName = sheetNameParam?.Value;
                string userToMatch = userParam?.Value;
                string expectedEmail = expectedEmailParam?.Value;

                if (!File.Exists(filePath))
                    return new UnknownFailedActionResult($"Excel file not found at path: {filePath}");

                List<string> errors = new List<string>();

                // Use user-provided mandatory columns or default
                List<string> mandatoryColumns = string.IsNullOrEmpty(mandatoryColsParam?.Value)
                    ? new List<string> { "Name", "Email", "Flow" }
                    : mandatoryColsParam.Value.Split(',').Select(x => x.Trim()).ToList();

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = string.IsNullOrEmpty(sheetName)
                        ? package.Workbook.Worksheets[0]
                        : package.Workbook.Worksheets[sheetName];

                    if (worksheet == null)
                        return new UnknownFailedActionResult("Specified worksheet not found.");

                    var columnMap = new Dictionary<string, int>();
                    int totalColumns = worksheet.Dimension.End.Column;
                    int totalRows = worksheet.Dimension.End.Row;

                    // Read headers
                    for (int col = 1; col <= totalColumns; col++)
                    {
                        string header = worksheet.Cells[1, col].Text.Trim();
                        if (!string.IsNullOrEmpty(header) && !columnMap.ContainsKey(header))
                            columnMap[header] = col;
                    }

                    // Verify mandatory columns
                    foreach (string col in mandatoryColumns)
                    {
                        if (!columnMap.ContainsKey(col))
                            errors.Add($"Missing mandatory column: '{col}'");
                    }

                    if (errors.Any())
                        return new UnknownFailedActionResult("Validation failed:\n" + string.Join("\n", errors));

                    // Dynamic content check
                    for (int row = 2; row <= totalRows; row++)
                    {
                        if (!string.IsNullOrEmpty(userToMatch) && !string.IsNullOrEmpty(expectedEmail))
                        {
                            string userValue = worksheet.Cells[row, columnMap["Name"]].Text.Trim();
                            string emailValue = worksheet.Cells[row, columnMap["Email"]].Text.Trim();

                            if (userValue == userToMatch && emailValue != expectedEmail)
                            {
                                errors.Add($"Row {row}: Name = '{userValue}' but Email = '{emailValue}' (Expected: '{expectedEmail}')");
                            }
                        }
                    }
                }

                if (errors.Any())
                    return new UnknownFailedActionResult("Validation issues:\n" + string.Join("\n", errors));

                return new PassedActionResult("Excel validation passed successfully.");
            }
            catch (Exception ex)
            {
                return new UnknownFailedActionResult($"Exception: {ex.Message}");
            }
        }
    }
}
