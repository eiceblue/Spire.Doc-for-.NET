Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AppendColumnChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim document As New Document()

			' Add a section to the document
			Dim section As Section = document.AddSection()

			' Add a paragraph to the section and append text to it
			section.AddParagraph().AppendText("Column chart.")

			' Add a new paragraph to the section
			Dim newPara As Paragraph = section.AddParagraph()

			' Append a column chart shape to the paragraph with specified width and height
			Dim shape As ShapeObject = newPara.AppendChart(ChartType.Column, 500, 300)

			' Get the chart object from the shape
			Dim chart As Chart = shape.Chart

			' Clear any existing series in the chart
			chart.Series.Clear()

			' Add a new series to the chart with data points for X values (categories) and Y values
			chart.Series.Add("Test Series", { "Word", "PDF", "Excel", "GoogleDocs", "Office" }, New Double() { 1900000, 850000, 2100000, 600000, 1500000 })

			' Set the number format for the Y-axis labels
			chart.AxisY.NumberFormat.FormatCode = "#,##0"

			' Save the document to a file in Docx format
			document.SaveToFile("AppendColumnChart.docx", FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launching the Word file.
			WordDocViewer("AppendColumnChart.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
