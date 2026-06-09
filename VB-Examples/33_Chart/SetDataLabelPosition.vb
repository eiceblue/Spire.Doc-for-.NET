Imports System
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields
Imports Spire.Doc.Fields.Shapes

Namespace SetDataLabelPosition
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new Word document instance
            Dim doc As Document = New Document()

            ' Add a new section to the document
            Dim section As Section = doc.AddSection()

            ' Add a paragraph with the text "Center" as a title/label
            section.AddParagraph().AppendText("Center")

            ' Add a new paragraph to hold the first chart
            Dim newPara As Spire.Doc.Documents.Paragraph = section.AddParagraph()

            ' Append a Pie chart to the paragraph and set its size (width: 500, height: 300)
            Dim shape As ShapeObject = newPara.AppendChart(ChartType.Pie, 500, 300)

            ' Get the Chart object from the created shape
            Dim chart As Chart = shape.Chart

            ' Enable data labels for the first data series in the pie chart
            chart.Series[0].HasDataLabels = True

            ' Configure the data labels to display the category name
            chart.Series[0].DataLabels.ShowCategoryName = True

            ' Configure the data labels to display the numeric value
            chart.Series[0].DataLabels.ShowValue = True

            ' Set the position of the data labels to the center of the pie slices
            chart.Series[0].DataLabels.Position = ChartDataLabelPosition.Center

            ' Add another paragraph with the text "Left" as a title/label
            section.AddParagraph().AppendText("Left")

            newPara = section.AddParagraph()

            ' Append a Bubble chart to the same paragraph and set its size (width: 500, height: 300)
            Dim shape2 As ShapeObject = newPara.AppendChart(ChartType.Bubble, 500, 300)

            ' Get the Chart object from the second shape
            Dim chart2 As Chart = shape2.Chart

            ' Enable data labels for the first data series in the bubble chart
            chart2.Series[0].HasDataLabels = True

            ' Configure the data labels to display the category name
            chart2.Series[0].DataLabels.ShowCategoryName = True

            ' Configure the data labels to display the numeric value
            chart2.Series[0].DataLabels.ShowValue = True

            ' Set the position of the data labels to the left side
            chart2.Series[0].DataLabels.Position = ChartDataLabelPosition.Left

            ' Define the output file name for saving the document
            Dim outputFile As String = "SetDataLabelPosition.docx"

            ' Save the document to the specified file in Docx format
            doc.SaveToFile(outputFile, FileFormat.Docx)

            ' Close the document and release associated resources
            doc.Close()

            ' Dispose of the document object to free up memory
            doc.Dispose()

            WordDocViewer(outputFile)
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

    End Class
End Namespace
