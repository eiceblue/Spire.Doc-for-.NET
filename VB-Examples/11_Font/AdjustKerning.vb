Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AdjustKerning
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new instance of the Document class
            Dim doc As Document = New Document()

            ' Add a new section to the document
            Dim section As Section = doc.AddSection()

            ' Create a list to store test data (text description and kerning value)
            Dim testData As List<Object[]> = New List<Object[]>()

            ' Add a test case for negative kerning
            testData.Add(New Object[] { "Negative Kerning (-1.0f)", -1.0f })

            ' Add a test case for zero kerning (disables kerning)
            testData.Add(New Object[] { "Zero Kerning (0.0f)", 0.0f })

            ' Add a test case for positive kerning
            testData.Add(New Object[] { "Positive Kerning (2.5f)", 2.5f })

            ' Add a test case for a large kerning value
            testData.Add(New Object[] { "Huge Kerning (1638.0f)", 1638.0f })

            ' Add a test case for a value exceeding the standard limit (1-1638)
            testData.Add(New Object[] { "Tiny Kerning (1639.0f)", 1639.0f })

            ' Loop through each test data item
            For Each object[] item in testData
                ' Extract the text description from the first column
                Dim text As String = (String)item[0]

                ' Extract the kerning value from the second column
                Dim kerningValue As Single = (Single)item[1]

                ' Add a new paragraph to the section
                Dim pragraph As Paragraph = section.AddParagraph()

                ' Append the text to the paragraph and get the text range
                Dim textRange As TextRange = pragraph.AppendText(text)

                ' Apply the specific kerning value to the character format
                textRange.CharacterFormat.Kerning = kerningValue
            Next

            ' Define the file name for the output document
            Dim result As String = "Adjust Kerning.docx"

            ' Save the document to a file in Docx format
            doc.SaveToFile(result, FileFormat.Docx)

            ' Close the document to release resources
            doc.Close()

            'Launching the Word file.
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
