Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AppendPieChart
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
			section.AddParagraph().AppendText("Pie chart.")

			' Add a new paragraph to the section
			Dim newPara As Paragraph = section.AddParagraph()

			' Append a pie chart shape to the paragraph with specified width and height
			Dim shape As ShapeObject = newPara.AppendChart(ChartType.Pie, 500, 300)
			Dim chart As Chart = shape.Chart

			' Add a series to the chart with categories (labels) and corresponding data values
			Dim series As ChartSeries = chart.Series.Add("Test Series", { "Word", "PDF", "Excel" }, { 2.7, 3.2, 0.8 })

			' Save the document to a file in Docx format
			document.SaveToFile("AppendPieChart.docx", FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("AppendPieChart.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
