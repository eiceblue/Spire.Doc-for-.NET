Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AppendChartLegend
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
			Dim document As New Document()

			'Load the file from disk.
			document.LoadFromFile("..\..\..\..\..\..\Data\ChartTemplate.docx")

			' Loop through all sections in the document
			For i As Integer = 0 To document.Sections.Count - 1
				' Loop through all paragraphs in the current section
				For j As Integer = 0 To document.Sections(i).Paragraphs.Count - 1
					' Get the current paragraph
					Dim paragraph = document.Sections(i).Paragraphs(j)

					' Loop through all child objects in the paragraph
					For Each obj As DocumentObject In paragraph.ChildObjects
						' Check if the object is a shape (e.g., chart, etc.)
						If TypeOf obj Is ShapeObject Then
							' Cast the object to a ShapeObject
							Dim shape = TryCast(obj, ShapeObject)

							' Get the chart from the shape
							Dim chart As Chart = shape.Chart

							' Call the method to add or update the chart legend
							AppendChartLegend(chart)
						End If
					Next obj
				Next j
			Next i


			document.SaveToFile("AppendChartLegend.docx", FileFormat.Docx2019)

			'Dispose the document
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("AppendChartLegend.docx")
		End Sub
		Public Sub AppendChartLegend(ByVal chart As Spire.Doc.Fields.Shapes.Charts.Chart)
			' Enable the legend display on the chart
			chart.Legend.Show = True

			' Set the position of the legend to the left side of the chart
			chart.Legend.Position = LegendPosition.Left

			' Disable overlay mode so the legend does not overlap with the chart plot area
			chart.Legend.Overlay = False

			' Set the font size of the legend text to 9 points
			chart.Legend.CharacterFormat.FontSize = 9

			' Set the text color of the legend labels to blue
			chart.Legend.CharacterFormat.TextColor = Color.Blue

			' Apply italic style to the legend text
			chart.Legend.CharacterFormat.Italic = True
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
