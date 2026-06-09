Imports System
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields

Namespace CombinedChart
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Initialize a new Document object
            Dim doc As Document = New Document()

            ' Add a new section to the document and create a paragraph within that section
            Dim paragraph As Paragraph = doc.AddSection().AddParagraph()

            ' Append a column chart (450x300 pixels) to the paragraph and retrieve the Chart object
            Dim chart As Chart = paragraph.AppendChart(ChartType.Column, 450, 300).Chart

            ' Change the chart type of the series named "Series 3" to a Line chart and enable secondary axis if applicable
            chart.ChangeSeriesType("Series 3", ChartSeriesType.Line, True)

            ' Define the output file name for the combined chart document
            Dim outputFile As String = "CombinedChart.docx"

            ' Save the document to the specified file in DOCX 2019 format
            doc.SaveToFile(outputFile, FileFormat.Docx2019)

            ' Close the document to release resources
            doc.Close()

            ' Dispose of the document object to free up memory
            doc.Dispose()

            'Launch the Word file.
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
