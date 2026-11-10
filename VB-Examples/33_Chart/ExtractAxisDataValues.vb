Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields
Imports System.Text
Imports System.IO

Namespace ExtractAxisDataValues
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object for the first document
			Dim doc As New Document()

			' Load the Word document from the specified relative file path
			doc.LoadFromFile("..\..\..\..\..\..\Data\ExtractAxisDataValues.docx")

			' Initialize a StringBuilder to store the extracted axis data values
			Dim stringBuilder As New StringBuilder()

			' Loop through each section in the loaded document
			For Each sec As Section In doc.Sections
				' Loop through each paragraph within the current section
				For Each paragraph As Paragraph In sec.Paragraphs
					' Iterate over all child objects contained in the paragraph
					For i As Integer = 0 To paragraph.ChildObjects.Count - 1
						' Get the current document object at index i
						Dim obj As DocumentObject = paragraph.ChildObjects(i)

						' Check if the current object is a ShapeObject (which can contain charts)
						If TypeOf obj Is ShapeObject Then
							' Cast the object to a ShapeObject
							Dim shape As ShapeObject = TryCast(obj, ShapeObject)

							' Retrieve the Chart object from the shape
							Dim chart As Chart = shape.Chart

							' Add a header line indicating the start of X-axis data extraction
							stringBuilder.AppendLine("Obtain X-axis data values:")

							' Loop through all the X-axis values in the chart
							For x As Integer = 0 To chart.XValues.Count - 1
								' Get the specific X-axis value at the current index
								Dim xVal As ChartValue = chart.XValues(x)

								' Append the string representation of the X-value to the StringBuilder, followed by a space
								stringBuilder.Append(xVal.StringValue & " ")
							Next x

							' Get the first data series from the chart (index 0)
							Dim series As ChartSeries = chart.Series(0)

							' Add a new line and a header for Y-axis data extraction
							stringBuilder.AppendLine(vbCrLf & "Obtain Y-axis data values:")

							' Iterate through all the Y-values in the selected data series
							For Each yVal As ChartValue In series.YValues
								' Append the numeric value of the Y-data point to the StringBuilder, followed by a space
								stringBuilder.Append(yVal.Value & " ")
							Next yVal
						End If
					Next i
				Next paragraph
			Next sec

			' Define the output file name/path for saving the extracted data
			Dim result As String = "ExtractAxisDataValues.txt"

			' Write the entire content of the StringBuilder to the specified text file
			File.WriteAllText(result, stringBuilder.ToString())

			' Close the document and release associated resources
			doc.Close()

			' Dispose of the document object to free up memory
			doc.Dispose()

			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
