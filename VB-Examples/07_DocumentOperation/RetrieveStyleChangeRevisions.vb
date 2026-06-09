Imports System
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Fields
Imports System.IO

Namespace RetrieveStyleChangeRevisions
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Initialize a new, empty Document object.
            Dim doc As Document = New Document()

            ' Load the existing Word document containing revision history from the specified file path.
            doc.LoadFromFile(@"..\..\..\..\..\..\..\Data\GetRevisions.docx")

            ' Retrieve the collection of all revision information (changes, comments, etc.) from the document.
            Dim revisionInfoCollection As RevisionInfoCollection = doc.GetRevisionInfos()

            ' Initialize a StringBuilder to efficiently construct the output text report.
            Dim stringBuilder As StringBuilder = New StringBuilder()

            ' Iterate through each revision item in the collected revision information.
            For Each revisionInfo As RevisionInfo In revisionInfoCollection
                ' Check if the current revision is specifically a formatting change (e.g., bold, color, font).
                If revisionInfo.RevisionType == RevisionType.FormatChange Then
                    ' Verify if the object affected by this revision is a TextRange (a segment of text).
                    If revisionInfo.OwnerObject is Spire.Doc.Fields.TextRange Then
                        ' Cast the owner object to a TextRange to access its specific properties.
                        Dim range As TextRange = (TextRange)revisionInfo.OwnerObject

                        ' Append the actual text content of the modified range to the report.
                        stringBuilder.AppendLine("TextRange:" + range.Text + "rn")

                        ' Switch the document view to the "Original" state to read pre-change formatting properties.
                        doc.RevisionsView = RevisionsView.Original

                        ' Append the original formatting details (Bold, Color, Highlight, Font, Underline) to the report.
                        stringBuilder.AppendLine("Original style£º" + "isBold£º" + range.CharacterFormat.Bold + ";" + "TextColor£º" + range.CharacterFormat.TextColor + "£»HighlightColor£º" + range.CharacterFormat.HighlightColor + "£»FontName£º" + range.CharacterFormat.FontName + "£»UnderlineStyle£º" + range.CharacterFormat.UnderlineStyle + "rn")

                        ' Switch the document view to the "Final" state to read post-change formatting properties.
                        doc.RevisionsView = RevisionsView.Final

                        ' Append the final formatting details to compare against the original state.
                        stringBuilder.AppendLine("Final style£º" + "isBold£º" + range.CharacterFormat.Bold + ";" + "TextColor£º" + range.CharacterFormat.TextColor + "£»HighlightColor£º" + range.CharacterFormat.HighlightColor + "£»FontName£º" + range.CharacterFormat.FontName + "£»UnderlineStyle£º" + range.CharacterFormat.UnderlineStyle + "rn")
                    End If
                End If
            Next

            ' Write the complete accumulated report string to a text file.
            File.WriteAllText("RetrieveStyleChangeRevisions.txt", stringBuilder.ToString())

            ' Close the document to release file resources and memory.
            doc.Close()
        End Sub
    End Class
End Namespace
