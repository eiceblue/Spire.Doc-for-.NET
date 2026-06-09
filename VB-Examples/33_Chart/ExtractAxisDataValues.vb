Imports System
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields
Imports System.Text
Imports System.IO

Namespace ExtractAxisDataValues
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new Document object for the first document
            Dim doc As Document = New Document()

            ' Load the Word document from the specified relative file path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractAxisDataValues.docx")

            ' Initialize a StringBuilder to store the extracted axis data values
            Dim stringBuilder As StringBuilder = New StringBuilder()

            ' Loop through each section in the loaded document
            For Each sec As Section In doc.Sections
                ' Loop through each paragraph within the current section
                For Each paragraph As Paragraph In sec.Paragraphs
                    ' Iterate over all child objects contained in the paragraph
                    Dim i As Integer = 0
                    While i < paragraph.ChildObjects.Count
                        ' Get the current document object at index i
                        Dim obj As DocumentObject = paragraph.ChildObjects[i]

                        ' Check if the current object is a ShapeObject (which can contain charts)
                        If obj is ShapeObject Then
                            ' Cast the object to a ShapeObject
                            Dim shape As ShapeObject = obj as ShapeObject

                            ' Retrieve the Chart object from the shape
                            Dim chart As Chart = shape.Chart

                            ' Add a header line indicating the start of X-axis data extraction
                            stringBuilder.AppendLine("Obtain X-axis data values:")

                            ' Loop through all the X-axis values in the chart
                            Dim x As Integer = 0
                            While x < chart.XValues.Count
                                ' Get the specific X-axis value at the current index
                                Dim xVal As ChartValue = chart.XValues[x]

                                ' Append the string representation of the X-value to the StringBuilder, followed by a space
                                stringBuilder.Append(xVal.StringValue + " ")
                            Next

                            ' Get the first data series from the chart (index 0)
                            Dim series As ChartSeries = chart.Series[0]

                            ' Add a new line and a header for Y-axis data extraction
                            stringBuilder.AppendLine("rnObtain Y-axis data values:")

                            ' Iterate through all the Y-values in the selected data series
                            For Each yVal As ChartValue In series.YValues
                                ' Append the numeric value of the Y-data point to the StringBuilder, followed by a space
                                stringBuilder.Append(yVal.Value + " ")
                            Next
                        End If
                    Next
                Next
            Next

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
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

    End Class
End Namespace
