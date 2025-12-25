Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AddCombinationChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document instance
			Dim document As Document = New Document()

			' Add a section to the document and then add a paragraph to that section
			Dim paragraph As Paragraph = document.AddSection().AddParagraph()

			' Append a chart of specified type and size to the paragraph, and get the Chart object
			Dim chart As Chart = paragraph.AppendChart(ChartType.Column, 450, 300).Chart

			' Modify 'Series 3' to a line chart and set it to display on the secondary axis
			chart.ChangeSeriesType("Series 3", ChartSeriesType.Line, True)

			' Define the file path and name for saving the document
			Dim filePath As String = "AddCombinationChart.docx"

			' Save the document to the specified file path in DOCX format
			document.SaveToFile(filePath, FileFormat.Docx)

			' Release resources used by the document
			document.Dispose()

			'Launch the Word file.
			WordDocViewer(filePath)

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
