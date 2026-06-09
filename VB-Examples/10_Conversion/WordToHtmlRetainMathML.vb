Imports System
Imports System.Windows.Forms
Imports Spire.Doc

Namespace WordToHtmlRetainMathML
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Initialize a new Document object
            Dim document As Document = New Document()

            ' Load an existing Word document from the specified relative file path
            document.LoadFromFile(@"..\..\..\..\..\..\Data\GetMathEquation.docx")

            ' Retrieve the HTML export options configuration object for the document
            Dim htmlExportOptions As HtmlExportOptions = document.HtmlExportOptions

            ' Configure the export to render Office math equations using MathML format
            htmlExportOptions.OfficeMathOutputMode = HtmlOfficeMathOutputMode.MathML

            ' Set the CSS stylesheet to be embedded internally within the generated HTML file
            htmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal

            ' Define the output file name for the converted HTML document
            Dim outputFile As String = "WordToHtmlRetainMathML.html"

            ' Save the document as an HTML file using the configured export options
            document.SaveToFile(outputFile, FileFormat.Html)

            ' Close the document to release file handles
            document.Close()

            ' Dispose of the document object to free up memory
            document.Dispose()

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
