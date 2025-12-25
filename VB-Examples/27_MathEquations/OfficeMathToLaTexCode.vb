Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.OMath

Namespace OfficeMathToLaTexCode
    Partial Public Class Form1
        Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            ' Load the existing Word document containing Office Math objects
            Dim document As Document = New Document("..\..\..\..\..\..\Data\OfficeMath.docx")

            ' Create a StringBuilder to accumulate the LaTeX code strings
            Dim stringBuilder As StringBuilder = New StringBuilder()

            ' Iterate through all sections in the document
            For Each section As Section In document.Sections
                ' Iterate through all paragraphs within the current section's body
                For Each par As Paragraph In section.Body.Paragraphs
                    ' Iterate through all child objects within the current paragraph
                    For Each obj As DocumentObject In par.ChildObjects
                        ' Attempt to cast the current object to an OfficeMath object
                        Dim officeMath As OfficeMath = TryCast(obj, OfficeMath)

                        ' If the cast fails (obj is not OfficeMath), skip to the next object
                        If officeMath Is Nothing Then Continue For

                        ' Convert the OfficeMath object to its LaTeX representation
                        Dim LaTexCode As String = officeMath.ToLaTexMathCode()

                        ' Append the LaTeX code to the StringBuilder, followed by a new line
                        stringBuilder.AppendLine(LaTexCode)
                    Next
                Next
            Next

            ' Define the name of the output text file
            Dim outputFile As String = "OfficeMathToLaTexCode.txt"

            ' Write the accumulated LaTeX code string to the output file
            File.WriteAllText(outputFile, stringBuilder.ToString())

            ' Dispose of the Document object to release resources
            document.Dispose()
        End Sub



    End Class
End Namespace
