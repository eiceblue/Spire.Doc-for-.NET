Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AppendChartTitle
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

							' Call the method to add or update the chart title
							AppendChartTitle(chart)
						End If
					Next obj
				Next j
			Next i


			document.SaveToFile("AppendChartTitle.docx",FileFormat.Docx2019)

			'Dispose the document
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("AppendChartTitle.docx")
		End Sub
		Public Sub AppendChartTitle(ByVal chart As Spire.Doc.Fields.Shapes.Charts.Chart)
			' Get the chart's title object
			Dim title As ChartTitle = chart.Title

			' Enable the display of the title
			title.Show = True

			' Disable overlay so the title does not overlap with the chart area
			title.Overlay = False

			' Set the text of the title
			title.Text = "My Chart"

			' Set font size of the title
			title.CharacterFormat.FontSize = 12

			' Set the title text to bold
			title.CharacterFormat.Bold = True

			' Set the text color to blue
			title.CharacterFormat.TextColor = Color.Blue

			' Enable right-to-left text formatting (if needed for language)
			title.CharacterFormat.Bidi = True

			' Apply italic style to the title text
			title.CharacterFormat.Italic = True

			' Set character spacing (tracking or kerning)
			title.CharacterFormat.CharacterSpacing = 2

			' Set underline color to red
			title.CharacterFormat.UnderlineColor = Color.Red

			' Set underline style to double line
			title.CharacterFormat.UnderlineStyle = UnderlineStyle.Double

			' Set font name
			title.CharacterFormat.FontName = "arial"

			' Enable all caps formatting
			title.CharacterFormat.AllCaps = True

			' Enable shadow effect on the text
			title.CharacterFormat.IsShadow = True

			' Set the position of the text baseline relative to normal 
			title.CharacterFormat.Position = 3
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
