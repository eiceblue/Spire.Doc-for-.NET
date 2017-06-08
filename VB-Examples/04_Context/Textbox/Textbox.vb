Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace InsertingTextbox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            'Open a blank word document as template
            Dim document_Renamed As New Document("..\..\..\..\..\..\Data\Blank.doc")

            InsertTextbox(document_Renamed.Sections(0))

            'Save doc file.
            document_Renamed.SaveToFile("Sample.doc", FileFormat.Doc)

            'Launching the MS Word file.
            WordDocViewer("Sample.doc")


        End Sub

        Private Sub InsertTextbox(ByVal section_Renamed As Section)
            Dim paragraph_Renamed As Paragraph = Nothing
            If section_Renamed.Paragraphs.Count > 0 Then
                paragraph_Renamed = section_Renamed.Paragraphs(0)
            Else
                paragraph_Renamed = section_Renamed.AddParagraph()
            End If
            paragraph_Renamed.AppendText("The sample demonstrates how to insert a textbox into a document.")
            paragraph_Renamed.ApplyStyle(BuiltinStyle.Heading2)

            paragraph_Renamed = section_Renamed.AddParagraph()
            paragraph_Renamed.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left
            Dim textBox As Spire.Doc.Fields.TextBox = paragraph_Renamed.AppendTextBox(50, 20)

        End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
