using MergeXlsx;

var root = Directory.GetCurrentDirectory();
var path = Path.Combine(root + "\\..\\..\\..\\");
var folder = Path.Combine(path, "Data");

//var files = new[]{ $"{folder}\\simple\\sample1.xlsx", $"{folder}\\simple\\sample2.xlsx" };
//var files = new[] { $"{folder}\\clients\\sample1.xlsx", $"{folder}\\clients\\sample2.xlsx" };

var files = new[] { $"{folder}\\sample1.xlsx", $"{folder}\\sample2.xlsx", $"{folder}\\sample3.xlsx" };


ClosedXmlMerger.Merge(files, $@"{path}\results\ClosedXml\merged-result.xlsx");
ClosedXmlMerger.MergeExcelFilesWithSameSheetNames([$"{folder}\\simple\\sample1.xlsx", $"{folder}\\simple\\sample2.xlsx"], $@"{path}\results\ClosedXml\merged-wordbook-result.xlsx");

OpenXmlMerger.Merge(files, $@"{path}\results\OpenXml\merged-result.xlsx");
