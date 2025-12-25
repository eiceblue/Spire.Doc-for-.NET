Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AppendChartDataLabel
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
							Dim series As ChartSeriesCollection = chart.Series
							Dim dataLabels As ChartDataLabelCollection = series(0).DataLabels
							series(0).HasDataLabels = True
							AppendChartDataLabel(dataLabels)
						End If
					Next obj
				Next j
			Next i

			document.SaveToFile("AppendChartDataLabel.docx",FileFormat.Docx2019)

			'Dispose the document
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("AppendChartDataLabel.docx")
		End Sub
		Public Sub AppendChartDataLabel(ByVal dataLabels As ChartDataLabelCollection)
			' Display the value (e.g., percentage or numerical value) on the data labels
			dataLabels.ShowValue = True

			' Display the category name (e.g., the label for each chart segment)
			dataLabels.ShowCategoryName = True

			' Display the series name (useful when multiple series are present)
			dataLabels.ShowSeriesName = True

			' Show leader lines connecting the data labels to the chart elements
			dataLabels.ShowLeaderLines = True

			' Set the separator between different label components (e.g., value and category)
			dataLabels.Separator = ";"

			' Set the number format for the displayed values (thousands separator and zero decimals)
			dataLabels.NumberFormat.FormatCode = "#,##0"

			' Set the font size of the data labels
			dataLabels.CharacterFormat.FontSize = 8

			' Make the text in the data labels bold
			dataLabels.CharacterFormat.Bold = True

			' Set the text color of the data labels to blue
			dataLabels.CharacterFormat.TextColor = Color.Blue

			' Set the border color of the characters in the data labels to blue
			dataLabels.CharacterFormat.Border.Color = Color.Blue

			' Enable right-to-left (RTL) text direction for languages like Arabic or Hebrew
			dataLabels.CharacterFormat.Bidi = True

			' Apply italic formatting to the text
			dataLabels.CharacterFormat.Italic = True

			' Set the underline color to red
			dataLabels.CharacterFormat.UnderlineColor = Color.Red

			' Set the underline style to double line
			dataLabels.CharacterFormat.UnderlineStyle = UnderlineStyle.Double

			' Set the font family for the data labels
			dataLabels.CharacterFormat.FontName = "Arial"

			' Display all text in uppercase letters
			dataLabels.CharacterFormat.AllCaps = True

			' Apply a shadow effect to the text
			dataLabels.CharacterFormat.IsShadow = True

			' Set the opacity (transparency) of the text effect (e.g., shadow or glow)
			dataLabels.CharacterFormat.TextEffectFormat.TextOpacity = 0.1
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
