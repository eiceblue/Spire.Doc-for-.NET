Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AppendScatterChart
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
			section.AddParagraph().AppendText("Scatter chart.")

			' Add a new paragraph to the section
			Dim newPara As Paragraph = section.AddParagraph()

			' Append a scatter chart shape to the paragraph with specified width and height
			Dim shape As ShapeObject = newPara.AppendChart(ChartType.Scatter, 450, 300)
			Dim chart As Chart = shape.Chart

			' Clear any existing series in the chart
			chart.Series.Clear()

			' Add a new series to the chart with data points for X and Y values
			chart.Series.Add("Scatter chart", { 1.0, 2.0, 3.0, 4.0, 5.0 }, { 1.0, 20.0, 40.0, 80.0, 160.0 })

			' Save the document to a file in Docx format
			document.SaveToFile("AppendScatterChart.docx", FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launching the Word file.
			WordDocViewer("AppendScatterChart.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
