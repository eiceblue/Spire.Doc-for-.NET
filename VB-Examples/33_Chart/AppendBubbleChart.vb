Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AppendBubbleChart
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
			section.AddParagraph().AppendText("Bubble chart.")

			' Add a new paragraph to the section
			Dim newPara As Paragraph = section.AddParagraph()

			' Append a bubble chart shape to the paragraph with specified width and height
			Dim shape As ShapeObject = newPara.AppendChart(ChartType.Bubble, 500, 300)

			' Get the chart object from the shape
			Dim chart As Chart = shape.Chart

			' Clear any existing series in the chart
			chart.Series.Clear()

			' Add a new series to the chart with data points for X, Y, and bubble size values
			Dim series As ChartSeries = chart.Series.Add("Test Series", { 2.9, 3.5, 1.1, 4.0, 4.0 }, { 1.9, 8.5, 2.1, 6.0, 1.5 }, { 9.0, 4.5, 2.5, 8.0, 5.0 })

			' Save the document to a file in Docx format
			document.SaveToFile("AppendBubbleChart.docx", FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("AppendBubbleChart.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
