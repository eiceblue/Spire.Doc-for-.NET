Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AppendSurface3DChart
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
			section.AddParagraph().AppendText("Surface3D chart.")

			' Add a new paragraph to the section
			Dim newPara As Paragraph = section.AddParagraph()

			' Append a Surface3D chart shape to the paragraph with specified width and height
			Dim shape As ShapeObject = newPara.AppendChart(ChartType.Surface3D, 500, 300)

			' Get the chart object from the shape
			Dim chart As Chart = shape.Chart

			' Clear any existing series in the chart
			chart.Series.Clear()

			' Set the title of the chart
			chart.Title.Text = "My chart"

			' Add multiple series to the chart with categories (X-axis values) and corresponding data values
			chart.Series.Add("Series 1", New String() { "Word", "PDF", "Excel", "GoogleDocs", "Office" }, New Double() { 1900000, 850000, 2100000, 600000, 1500000 })

			chart.Series.Add("Series 2", New String() { "Word", "PDF", "Excel", "GoogleDocs", "Office" }, New Double() { 900000, 50000, 1100000, 400000, 2500000 })

			chart.Series.Add("Series 3", New String() { "Word", "PDF", "Excel", "GoogleDocs", "Office" }, New Double() { 500000, 820000, 1500000, 400000, 100000 })

			' Save the document to a file in Docx format
			document.SaveToFile("AppendSurface3DChart.docx", FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launching the Word file.
			WordDocViewer("AppendSurface3DChart.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
