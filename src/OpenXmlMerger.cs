using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MergeXlsx;

public static class OpenXmlMerger
{
    public static void Merge(string[] inputFiles, string outPath)
    {
        var folder = outPath.Substring(0, outPath.LastIndexOf('\\'));
        if (!Directory.Exists(folder))
        {
            Directory.CreateDirectory(folder);
        }

        using SpreadsheetDocument destDoc = SpreadsheetDocument.Create(outPath, SpreadsheetDocumentType.Workbook);
        WorkbookPart destWorkbookPart = destDoc.WorkbookPart ?? destDoc.AddWorkbookPart();
        destWorkbookPart.Workbook = new Workbook();

        var sharedStringCount = 0;

        foreach (var src in inputFiles)
        {
            using SpreadsheetDocument sourceDoc = SpreadsheetDocument.Open(src, false);


            foreach (var sourceSheet in sourceDoc.WorkbookPart.Workbook.Descendants<Sheet>())
            {
                var sourceWorksheetPart = (WorksheetPart)sourceDoc.WorkbookPart.GetPartById(sourceSheet.Id);

                var currentWorksheetPart = destWorkbookPart.AddPart(sourceWorksheetPart);

                var sheets = destWorkbookPart.Workbook.GetFirstChild<Sheets>() ??
                             destWorkbookPart.Workbook.AppendChild(new Sheets());

                var newSheet = new Sheet()
                {
                    Id = destWorkbookPart.GetIdOfPart(currentWorksheetPart),
                    SheetId = (uint)(sheets.ChildElements.Count + 1),
                    Name = GetUniqueSheetName(destWorkbookPart, sourceSheet.Name)
                };

                sheets.AppendChild(newSheet);

                CopyRelatedParts(sourceDoc.WorkbookPart, destWorkbookPart, currentWorksheetPart, ref sharedStringCount);

                //todo process images and other resources
            }

        }

        destWorkbookPart.Workbook.Save();
    }

    static void CopyRelatedParts(WorkbookPart sourceWorkbookPart, WorkbookPart destWorkbookPart, WorksheetPart currentWorksheet, ref int sharedStringCount)
    {
        if (sourceWorkbookPart.SharedStringTablePart != null)
        {
            if (destWorkbookPart.SharedStringTablePart == null)
            {
                destWorkbookPart.AddNewPart<SharedStringTablePart>();
                destWorkbookPart.SharedStringTablePart.SharedStringTable = new SharedStringTable();
            }

            var count = 0;
            foreach (var item in sourceWorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
            {
                destWorkbookPart.SharedStringTablePart.SharedStringTable.AppendChild((SharedStringItem)item.Clone());
                count++;
            }

            if (sharedStringCount == 0)
            {
                sharedStringCount = count;
                return;
            }

            foreach (var cell in currentWorksheet.Worksheet.Descendants<Cell>())
            {
                if (cell.DataType == null || cell.DataType.Value != CellValues.SharedString) continue;

                cell.CellValue.Text = sharedStringCount.ToString();
                sharedStringCount++;
            }
        }

        //todo formulas, styles, etc..
    }

    private static string GetUniqueSheetName(WorkbookPart workbookPart, string desiredName)
    {
        var counter = 1;
        while (workbookPart.Workbook.Descendants<Sheet>().Any(s => s.Name == $"{desiredName}_{counter++}")) { }

        return $"{desiredName}_{counter}";
    }
}