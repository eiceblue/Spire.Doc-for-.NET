Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AppendBarChart
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
			section.AddParagraph().AppendText("Bar chart.")

			' Add a new paragraph to the section
			Dim newPara As Paragraph = section.AddParagraph()

			' Append a bar chart shape to the paragraph with specified width and height
			Dim chartShape As ShapeObject = newPara.AppendChart(ChartType.Bar, 400, 300)
			Dim chart As Chart = chartShape.Chart

			' Get the title of the chart
			Dim title As ChartTitle = chart.Title

			' Set the text of the chart title
			title.Text = "My Chart"

			' Show the chart title
			title.Show = True

			' Overlay the chart title on top of the chart
			title.Overlay = True

			' Save the document to a file in Docx format
			document.SaveToFile("AppendBarChart.docx", FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("AppendBarChart.docx")

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
