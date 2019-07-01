Imports DevExpress.Spreadsheet
Imports DevExpress.Web.ASPxSpreadsheet
Imports System
Imports DevExpress.Docs
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls

Namespace DXWebApplication1
	Partial Public Class WebForm1
		Inherits System.Web.UI.Page

		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
			ASPxSpreadsheet1.Open(MapPath("~/Docs/testDocument1.xlsx"))
		End Sub

		Protected Sub ASPxSpreadsheet1_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.CallbackEventArgsBase)
			Dim spreadSheet As ASPxSpreadsheet = TryCast(sender, ASPxSpreadsheet)
			Dim workbook As IWorkbook = spreadSheet.Document
			Dim worksheet As Worksheet = workbook.Worksheets(0)
			Select Case e.Parameter
				Case "applyFormatting"
					Dim priceRange As Range = worksheet.Range("C2:C15")
					Dim rangeFormatting As Formatting = priceRange.BeginUpdateFormatting()
					rangeFormatting.Font.Color = Color.SandyBrown
					rangeFormatting.Font.FontStyle = SpreadsheetFontStyle.Bold
					rangeFormatting.Fill.BackgroundColor = Color.PaleGoldenrod
					rangeFormatting.NumberFormat = "$0.0#"

					rangeFormatting.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
					rangeFormatting.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
					priceRange.EndUpdateFormatting(rangeFormatting)
				Case "insertLink"
					worksheet.Columns("G").WidthInPixels = 180
					Dim cell1 As Cell = worksheet.Cells("G4")
					cell1.Fill.BackgroundColor = Color.WhiteSmoke
					worksheet.Hyperlinks.Add(cell1, "https://documentation.devexpress.com/OfficeFileAPI/14912/Spreadsheet-Document-API", True, "Spreadsheet Document API")
				Case "drawBorders"
					Dim tableRange As Range = worksheet.Range("A2:E16")
					tableRange.Borders.SetAllBorders(Color.RosyBrown, BorderLineStyle.Hair)
				Case "showTotal"
					Dim cell2 As Cell = worksheet.Cells("E16")
					cell2.Formula = "=SUBTOTAL(9,E2:E15)"
					Dim cell3 As Cell = worksheet.Cells("A16")
					cell3.Formula = "SUBTOTAL(103,A2:A15)"
					Dim cell4 As Cell = worksheet.Cells("D16")
					cell4.Value = "Total amount"
			End Select

		End Sub
	End Class
End Namespace