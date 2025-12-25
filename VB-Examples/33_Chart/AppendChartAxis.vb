Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace AppendChartAxis
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

							' Call the method to add or update the chart axis
							AppendChartAxis(chart)
						End If
					Next obj
				Next j
			Next i


			document.SaveToFile("AppendChartAxis.docx", FileFormat.Docx2019)

			'Dispose the document
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("AppendChartAxis.docx")
		End Sub
		Public Sub AppendChartAxis(ByVal chart As Spire.Doc.Fields.Shapes.Charts.Chart)
			For i As Integer = 0 To chart.Axes.Count - 1
				If i = 0 Then
					chart.Axes(i).CategoryType = AxisCategoryType.Category
					chart.Axes(i).Bounds.Maximum = New AxisBound(5)
					chart.Axes(i).Bounds.Minimum = New AxisBound(0)
					chart.Axes(i).Units.Major = 1
					chart.Axes(i).Units.MajorTimeUnit = 0
					chart.Axes(i).Units.Minor = 1
					chart.Axes(i).Units.MinorTimeUnit = AxisTimeUnit.Days
					chart.Axes(i).HasMajorGridlines = False
					chart.Axes(i).HasMinorGridlines = True
					chart.Axes(i).Labels.IsAutoSpacing = False
					chart.Axes(i).Labels.Spacing = 1
					chart.Axes(i).Labels.Offset = 1
					chart.Axes(i).Labels.Position = AxisTickLabelPosition.Low
					chart.Axes(i).ReverseOrder = True
					chart.Axes(i).Title.Text = "x-axis"
					chart.Axes(i).Title.Show = True
					chart.Axes(i).Title.Overlay = True
				ElseIf i = 1 Then
					chart.Axes(i).CategoryType = 0
					chart.Axes(i).Units.IsMajorAuto = True
					chart.Axes(i).Units.IsMinorAuto = True
					chart.Axes(i).Bounds.LogBase = 10
					chart.Axes(i).HasMajorGridlines = True
					chart.Axes(i).HasMinorGridlines = False
					chart.Axes(i).ReverseOrder = False
					chart.Axes(i).Labels.IsAutoSpacing = True
					chart.Axes(i).Title.Text = "y-axis"
					chart.Axes(i).Title.Show = True
					chart.Axes(i).Title.Overlay = True
				Else
					chart.Axes(i).Title.Text = "z-axis"
					chart.Axes(i).Title.Show = True
					chart.Axes(i).Title.Overlay = False
				End If
				chart.Axes(i).Labels.Alignment = LabelAlignment.Left
				chart.Axes(i).Units.BaseTimeUnit = 0
				chart.Axes(i).AxisBetweenCategories = True
				chart.Axes(i).DisplayUnits.CustomUnit = 1
				chart.Axes(i).DisplayUnits.Unit = AxisBuiltInUnit.Custom
				chart.Axes(i).DisplayUnits.ShowLabel = True
				chart.Axes(i).TickMarks.Spacing = 1
				chart.Axes(i).TickMarks.Major = 0
				chart.Axes(i).TickMarks.Minor = AxisTickMark.Inside
				chart.Axes(i).Title.GetCharacterFormat().FontSize = 8
				chart.Axes(i).Title.GetCharacterFormat().TextColor = Color.Red
				chart.Axes(i).Title.GetCharacterFormat().Bold = True
			Next i
		End Sub
		 Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		 End Sub

	End Class
End Namespace
