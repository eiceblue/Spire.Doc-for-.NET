Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AppendLineChart
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
			section.AddParagraph().AppendText("Line chart.")

			' Add a new paragraph to the section
			Dim newPara As Paragraph = section.AddParagraph()

			' Append a line chart shape to the paragraph with specified width and height
			Dim shape As ShapeObject = newPara.AppendChart(ChartType.Line, 500, 300)

			' Get the chart object from the shape
			Dim chart As Chart = shape.Chart

			' Get the title of the chart
			Dim title As ChartTitle = chart.Title

			' Set the text of the chart title
			title.Text = "My Chart"

			' Clear any existing series in the chart
			Dim seriesColl As ChartSeriesCollection = chart.Series
			seriesColl.Clear()

			' Define categories (X-axis values)
			Dim categories() As String = { "C1", "C2", "C3", "C4", "C5", "C6" }

			' Add two series to the chart with specified categories and Y-axis values
			seriesColl.Add("AW Series 1", categories, New Double() { 1, 2, 2.5, 4, 5, 6 })
			seriesColl.Add("AW Series 2", categories, New Double() { 2, 3, 3.5, 6, 6.5, 7 })

			' Save the document to a file in Docx format
			document.SaveToFile("AppendLineChart.docx", FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launching the Word file.
			WordDocViewer("AppendLineChart.docx")

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
