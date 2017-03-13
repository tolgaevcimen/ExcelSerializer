using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.Linq;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Reflection;
using System.Collections;
using ExcelSerializer.CellWriters;

namespace ExcelSerializer
{
    public class ExcelSerializer
    {
        public ExcelPackage Package { get; set; }
        public ExcelWorksheet Sheet { get; set; }

        public string SheetName { get; set; }
        public string LogTitle { get; set; }
        public string Author { get; set; }
        public string Title { get; set; }

        public ExcelSerializer ()
        {
            SheetName = "Sheet1";
            LogTitle = "LINK CREATION REPORT";

            Author = "AçıkBulut";
            Title = "Report";
        }

        public void Serialize ( object objectToSerialize, string path )
        {
            var sheetName = "Sheet1";

            var ve = new ValueExtractor();
            ve.Extract(objectToSerialize);

            using ( Package = new ExcelPackage() )
            {
                Package.Workbook.Properties.Author = Author;
                Package.Workbook.Properties.Title = Title;

                Sheet = Package.Workbook.Worksheets.Add(sheetName);

                GetHeader();

                var dvw = new DirectValueWriter(ve.DirectValues, Sheet, 2, 1);
                dvw.Run();

                var ow = new ObjectWriter(ve.Objects, Sheet, 4, 1);
                ow.Run();

                Package.SaveAs(new System.IO.FileInfo(path));

                ////Create excel headers
                //Sheet.Cells[2, 1].Value = "Owner";
                //Sheet.Cells[2, 2].Value = "Link Id";
                //Sheet.Cells[2, 3].Value = "Link Name";
                //Sheet.Cells[2, 4].Value = "Date";

                //sheet.Row(2).Style.Font.Bold = true;
                //var rowIndex = 3;
                //foreach ( var downloadLink in downloadLinks )
                //{
                //    var col = 1;
                //    sheet.Cells[rowIndex, col++].Value = downloadLink.OwnerLogonName;
                //    sheet.Cells[rowIndex, col++].Value = downloadLink.Id;
                //    sheet.Cells[rowIndex, col++].Value = downloadLink.LinkName;
                //    sheet.Cells[rowIndex, col++].Value = downloadLink.CreateDate;

                //    rowIndex++;
                //    sheet.Cells["B" + rowIndex + ":K" + rowIndex].Merge = true;

                //    sheet.Cells[rowIndex, 1].Value = "Paths";
                //    sheet.Cells[rowIndex, 1].Style.Font.Bold = true;
                //    sheet.Cells[rowIndex, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                //    var paths = JsonConvert.DeserializeObject<List<string>>(downloadLink.Recipe);
                //    sheet.Cells[rowIndex, 2].Value = string.Join("\r\n", paths);
                //    sheet.Row(rowIndex).Height *= paths.Count;
                //    sheet.Row(rowIndex).Style.WrapText = true;

                //    rowIndex++;
                //    sheet.Cells["B" + rowIndex + ":K" + rowIndex].Merge = true;

                //    sheet.Cells[rowIndex, 1].Value = "Receivers";
                //    sheet.Cells[rowIndex, 1].Style.Font.Bold = true;

                //    var receivers = JsonConvert.DeserializeObject<List<user>>(downloadLink.SentUsersByDesktopAgent);
                //    sheet.Cells[rowIndex, 2].Value = string.Join("\r\n", receivers.Select(r => string.Format("{0} - {1}({2})", r.isGroup ? "group" : "user", r.name, r.id)));
                //    sheet.Row(rowIndex).Height *= receivers.Count;
                //    sheet.Row(rowIndex).Style.WrapText = true;

                //    rowIndex++;
                //}

                //sheet.Column(1).AutoFit();
                //sheet.Column(2).AutoFit();
                //sheet.Column(3).AutoFit();
                //sheet.Column(4).AutoFit();

                //sheet.Column(4).Style.Numberformat.Format = "dd/mm/yyyy";

                //ReturnFile(fileName, Package, sheet);
            }
        }

        private void AutoFitColumns ( ExcelWorksheet sheet )
        {
            var columnCnt = sheet.Dimension.End.Column;
            for ( int i = 1; i <= columnCnt; i++ )
            {
                sheet.Column(i).AutoFit();
            }
        }

        private void GetHeader ()
        {
            var first_row = Sheet.Cells[1, 1];
            Sheet.Cells["A1:K1"].Merge = true;

            first_row.Value = LogTitle;

            first_row.Style.Font.Bold = true;
            first_row.Style.Font.Size = 16;
        }

        public static string GetRange ( int row1, int column1, int row2, int column2 )
        {
            var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXY";

            var y1 = alphabet[column1 - 1].ToString();
            var y2 = alphabet[column2 - 1].ToString();

            return y1 + row1.ToString() + ":" + y2 + row2.ToString();
        }
    }
}
