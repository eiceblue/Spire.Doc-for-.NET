Imports System
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Fields
Imports System.IO
Imports Spire.Doc.Documents
Imports Spire.Doc.Settings

Namespace CompatibilityOptions
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Initialize a new Document object
            Dim doc As Document = New Document()

            ' Add a new section to the document
            Dim section As Section = doc.AddSection()

            ' Add a new paragraph to the section
            Dim paragraph As Paragraph = section.AddParagraph()

            ' Define a string containing a label and trailing spaces for the underline effect
            Dim blanks As String = "(6)                  "

            ' Append the text string to the paragraph and get the TextRange object
            Dim tr As TextRange = paragraph.AppendText(blanks)

            ' Set the underline style of the text to Single
            tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single

            ' Enable compatibility option: Include trailing spaces in underlined text
            doc.CompatibilityOptions.UlTrailSpace = True

            ' Enable compatibility option: Adjust line height specifically for tables
            doc.CompatibilityOptions.AdjustLineHeightInTable = True

            ' Enable compatibility option: Reserve space for underline characters to prevent clipping
            doc.CompatibilityOptions.SpaceForUL = True

            ' Enable compatibility option: Apply complex script breaking rules for line breaks
            doc.CompatibilityOptions.ApplyBreakingRules = True

            ' Disable compatibility option: Allow expansion of lines ending with Shift+Return (manual line break)
            doc.CompatibilityOptions.DoNotExpandShiftReturn = False

            ' Disable compatibility option: Do not override font size and justification defined in table styles
            doc.CompatibilityOptions.OverrideTableStyleFontSizeAndJustification = False

            ' Enable compatibility option: Prevent automatic fitting of tables that have fixed width constraints
            doc.CompatibilityOptions.DoNotAutofitConstrainedTables = True

            ' Optimize the document's compatibility settings specifically for Word 2016
            doc.CompatibilityOptions.OptimizeForWordVersion(WordVersion.Word2016)

            ' Define the output file name
            Dim outputFile As String = "CompatibilityOptions.docx"

            ' Save the document to the specified file in DOCX 2019 format
            doc.SaveToFile(outputFile, FileFormat.Docx2019)

            ' Close the document to release file handles
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
