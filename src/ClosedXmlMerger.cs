using ClosedXML.Excel;

namespace MergeXlsx;

/// <inheritdoc cref="https://docs.closedxml.io/en/latest/"/>
public static class ClosedXmlMerger
{
    public static void Merge(string[] inputFiles, string outPath)
    {
        MergeExcelFiles(inputFiles, outPath);
    }

    private static void MergeExcelFiles(string[] inputFiles, string outputFile)
    {
        var sheetNames = new HashSet<string>();

        using var workbook = new XLWorkbook();
        foreach (var file in inputFiles)
        {
            using var sourceWorkbook = new XLWorkbook(file);
            foreach (var worksheet in sourceWorkbook.Worksheets)
            {
                worksheet.CopyTo(workbook, GetName(workbook.Worksheets, sheetNames, worksheet.Name));
            }
        }

        workbook.SaveAs(outputFile);

    }

    private static string GetName(IXLWorksheets sourceWorksheets, HashSet<string> existNames, string name)
    {
        int i = 1;
        string curName = name;
        while (true)
        {
            if (sourceWorksheets.Any(x => x.Name == curName) || existNames.Any(x => x == curName))
            {
                i++;
                curName = $"{name}-{i}";
                continue;
            }

            return SetName(curName);
        }

        string SetName(string name)
        {
            existNames.Add(name);
            return name;
        }
    }

    /// <summary>
    /// Combining sheets from different files with the same name into one Workbook
    /// </summary>
    public static void MergeExcelFilesWithSameSheetNames(string[] sourceFilePaths, string outputFilePath)
    {
        using var targetWorkbook = new XLWorkbook();
        var sheetsData = new Dictionary<string, IXLWorksheet>();

        foreach (var filePath in sourceFilePaths)
        {
            using var sourceWorkbook = new XLWorkbook(filePath);
            foreach (var sourceSheet in sourceWorkbook.Worksheets)
            {
                if (!sheetsData.TryGetValue(sourceSheet.Name, out var targetSheet))
                {
                    targetSheet = targetWorkbook.Worksheets.Add(sourceSheet.Name);
                    sheetsData[sourceSheet.Name] = targetSheet;

                    var headerRow = sourceSheet.FirstRowUsed();
                    if (headerRow != null)
                    {
                        var headerRange = headerRow.FirstCell();
                        headerRange.CopyTo(targetSheet.Cell(1, 1));
                    }
                }

                var targetLastRow = targetSheet.LastRowUsed()?.RowNumber() ?? 0;

                var dataRange = sourceSheet.RangeUsed().Rows()
                    .Select(row => row.Cells());

                foreach (var cells in dataRange)
                {
                    targetLastRow++;

                    var cellIndex = 1;
                    foreach (var cell in cells)
                    {
                        cell.CopyTo(targetSheet.Cell(targetLastRow, cellIndex));
                        cellIndex++;
                    }
                }
            }
        }

        foreach (var worksheet in targetWorkbook.Worksheets)
        {
            worksheet.Columns().AdjustToContents();

            if (worksheet.FirstRowUsed() != null)
            {
                worksheet.RangeUsed().SetAutoFilter();
            }
        }

        targetWorkbook.SaveAs(outputFilePath);
    }
}