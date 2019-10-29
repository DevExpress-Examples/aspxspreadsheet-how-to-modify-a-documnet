using DevExpress.Spreadsheet;
using DevExpress.Web.ASPxSpreadsheet;
using System;
using DevExpress.Docs;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DXWebApplication1 {
    public partial class WebForm1 : System.Web.UI.Page {
        protected void Page_Load(object sender, EventArgs e) {
            ASPxSpreadsheet1.Open(MapPath("~/Docs/testDocument1.xlsx"));
        }

        protected void ASPxSpreadsheet1_Callback(object sender, DevExpress.Web.CallbackEventArgsBase e) {
            ASPxSpreadsheet spreadSheet = sender as ASPxSpreadsheet;
            IWorkbook workbook = spreadSheet.Document;
            Worksheet worksheet = workbook.Worksheets[0];
            switch (e.Parameter) {
                case "applyFormatting":
                    CellRange priceRange = worksheet.Range["C2:C15"];
                    Formatting rangeFormatting = priceRange.BeginUpdateFormatting();
                    rangeFormatting.Font.Color = Color.SandyBrown;
                    rangeFormatting.Font.FontStyle = SpreadsheetFontStyle.Bold;
                    rangeFormatting.Fill.BackgroundColor = Color.PaleGoldenrod;
                    rangeFormatting.NumberFormat = "$0.0#";

                    rangeFormatting.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                    rangeFormatting.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                    priceRange.EndUpdateFormatting(rangeFormatting);
                    break;
                case "insertLink":
                    worksheet.Columns["G"].WidthInPixels = 180;
                    CellRange cell1 = worksheet.Cells["G4"];
                    cell1.Fill.BackgroundColor = Color.WhiteSmoke;
                    worksheet.Hyperlinks.Add(cell1, "https://documentation.devexpress.com/OfficeFileAPI/14912/Spreadsheet-Document-API", true, "Spreadsheet Document API");
                    break;
                case "drawBorders":
                    CellRange tableRange = worksheet.Range["A2:E16"];
                    tableRange.Borders.SetAllBorders(Color.RosyBrown, BorderLineStyle.Hair);
                    break;
                case "showTotal":
                    CellRange cell2 = worksheet.Cells["E16"];
                    cell2.Formula = "=SUBTOTAL(9,E2:E15)";
                    CellRange cell3 = worksheet.Cells["A16"];
                    cell3.Formula = "SUBTOTAL(103,A2:A15)";
                    CellRange cell4 = worksheet.Cells["D16"];
                    cell4.Value = "Total amount";
                    break;
            }

        }
    }
}