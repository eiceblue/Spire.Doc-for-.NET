Imports System
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Fields
Imports Spire.Doc.Fields.Shapes

Namespace SaveChartToTemplate
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new instance of the Word document
            Dim doc As Document = New Document()

            ' Add a new section to the document
            Dim section As Section = doc.AddSection()

            ' Add a new paragraph to a newly created section
            Dim paragraph As Paragraph = section.AddParagraph()

            ' Append a column chart to the paragraph and retrieve the Chart object
            Dim chart As Chart = ((Shape)paragraph.AppendChart(ChartType.Column, 400, 300)).Chart

            ' Save the chart as a template file (.crtx)
            chart.SaveAsTemplate("SaveChartToTemplate.crtx")

            ' Close the document and release associated resources
            doc.Close()

            ' Dispose of the document object to free up memory
            doc.Dispose()

            Me.Close()
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

    End Class
End Namespace
