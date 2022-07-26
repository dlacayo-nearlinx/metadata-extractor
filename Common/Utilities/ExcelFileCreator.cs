using System.Collections.Generic;
using MetadataExtractor.Common.Models;
using SpreadsheetLight;

namespace MetadataExtractor.Common.Utilities
{
    public static class ExcelFileCreator
    {
        public static void ExportDataToExcel(List<FileMetadata> files, string filePath)
        {
            // generate excel file
            var sl = new SLDocument();
            var currentRow = 1;
            //var currentColumn = 1;

            sl.SetCellValue(currentRow, 1, "Name");
            sl.SetCellValue(currentRow, 2, "Name");
            sl.SetCellValue(currentRow, 3, "Folder");
            sl.SetCellValue(currentRow, 4, "Path");
            sl.SetCellValue(currentRow, 5, "Created");
            sl.SetCellValue(currentRow, 6, "Modified");
            sl.SetCellValue(currentRow, 7, "Size");
            sl.SetCellValue(currentRow, 8, "Title");
            sl.SetCellValue(currentRow, 9, "Subject");
            sl.SetCellValue(currentRow, 10, "Author");
            sl.SetCellValue(currentRow, 11, "Category");
            sl.SetCellValue(currentRow, 12, "Comments");
            sl.SetCellValue(currentRow, 13, "Keywords");

            currentRow++;
            foreach (var md in files)
            {
                sl.SetCellValue(currentRow, 1, md.Name);
                sl.SetCellValue(currentRow, 2, md.Extension);
                sl.SetCellValue(currentRow, 3, md.Folder);
                sl.SetCellValue(currentRow, 4, md.Path);
                sl.SetCellValue(currentRow, 5, md.Created);
                sl.SetCellValue(currentRow, 6, md.Modified);
                sl.SetCellValue(currentRow, 7, md.Size);
                sl.SetCellValue(currentRow, 8, md.Title);
                sl.SetCellValue(currentRow, 9, md.Subject);
                sl.SetCellValue(currentRow, 10, md.Author);
                sl.SetCellValue(currentRow, 11, md.Category);
                sl.SetCellValue(currentRow, 12, md.Comments);
                sl.SetCellValue(currentRow, 13, md.Keywords);
                currentRow++;
            }

            sl.SaveAs(filePath);
        }
    }
}
