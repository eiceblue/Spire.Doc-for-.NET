Imports System
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents.Comparison

Namespace CompareDocumentsWithOptions
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new Document object for the first document
			Dim doc1 As Document = New Document()

			' Load the first document from the specified file path
			doc1.LoadFromFile(@"..\..\..\..\..\..\Data\SupportDocumentCompare1.docx")

			' Create a new Document object for the second document
			Dim doc2 As Document = New Document()

			' Load the second document from the specified file path
			doc2.LoadFromFile(@"..\..\..\..\..\..\Data\SupportDocumentCompare2.docx")

            ' Create CompareOptions object and set IgnoreFormatting property to true
            Dim options As CompareOptions = New CompareOptions()
            options.CompareMoves = False
            options.IgnoreCaseChanges = False
            options.IgnoreComments = True
            options.IgnoreFields = True
            options.IgnoreFootnotes = True
            options.IgnoreTables = True
            options.IgnoreTextboxes = True

            ' Compare the contents of the two documents with specified options and mark differences using "E-iceblue" as the author name and current date and time
            doc1.Compare(doc2, "E-iceblue", System.DateTime.Now, options)

			' Specify the output file name for the compared result
			Dim result As String = "CompareDocumentsWithOptions_result.docx"

			' Save the compared result to the specified file path in Docx2013 format
			doc1.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the Document objects to release resources
			doc1.Dispose()
			doc2.Dispose()
			
            'View the document
            FileViewer(result)
        End Sub

        Private Sub FileViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub
    End Class
End Namespace
