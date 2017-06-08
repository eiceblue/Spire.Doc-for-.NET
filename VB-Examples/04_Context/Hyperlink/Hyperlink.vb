Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace Hyperlink
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            'Open a blank word document as template
            Dim document_Renamed As New Document("..\..\..\..\..\..\Data\Blank.doc")

            InsertHyberlink(document_Renamed.Sections(0))

            'Save doc file.
            document_Renamed.SaveToFile("Sample.doc", FileFormat.Doc)

            'Launching the MS Word file.
            WordDocViewer("Sample.doc")


        End Sub

        Private Sub InsertHyberlink(ByVal section_Renamed As Section)
            Dim paragraph_Renamed As Paragraph = Nothing
            If section_Renamed.Paragraphs.Count > 0 Then
                paragraph_Renamed = section_Renamed.Paragraphs(0)
            Else
                paragraph_Renamed = section_Renamed.AddParagraph()
            End If
            paragraph_Renamed.AppendText("Spire.XLS for .NET " & vbCrLf & " e-iceblue company Ltd. 2002-2010 All rights reserverd")
            paragraph_Renamed.ApplyStyle(BuiltinStyle.Heading2)

            paragraph_Renamed = section_Renamed.AddParagraph()
            paragraph_Renamed.AppendText("Home page")
            paragraph_Renamed.ApplyStyle(BuiltinStyle.Heading2)
            paragraph_Renamed = section_Renamed.AddParagraph()
            paragraph_Renamed.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)

            paragraph_Renamed = section_Renamed.AddParagraph()
            paragraph_Renamed.AppendText("Contact US")
            paragraph_Renamed.ApplyStyle(BuiltinStyle.Heading2)
            paragraph_Renamed = section_Renamed.AddParagraph()
            paragraph_Renamed.AppendHyperlink("mailto:support@e-iceblue.com", "support@e-iceblue.com", HyperlinkType.EMailLink)
        End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
