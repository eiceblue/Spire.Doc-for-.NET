Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace DocumentSetup
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            'Open a blank word document as template
            Dim document As New Document("..\..\..\..\..\..\Data\Blank.doc")

            document.BuiltinDocumentProperties.Title = "Document Demo Document"
            document.BuiltinDocumentProperties.Subject = "demo"
            document.BuiltinDocumentProperties.Author = "James"
            document.BuiltinDocumentProperties.Company = "e-iceblue"
            document.BuiltinDocumentProperties.Manager = "Jakson"
            document.BuiltinDocumentProperties.Category = "Doc Demos"
            document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo"
            document.BuiltinDocumentProperties.Comments = "This document is just a demo."

            Dim section As Section = document.Sections(0)
            section.PageSetup.Margins.Top = 72.0F
            section.PageSetup.Margins.Bottom = 72.0F
            section.PageSetup.Margins.Left = 89.85F
            section.PageSetup.Margins.Right = 89.85F

            Dim p1 As String _
                = "Microsoft Word is a word processor designed by Microsoft. " _
                + "It was first released in 1983 under the name Multi-Tool Word for Xenix systems. " _
                + "Subsequent versions were later written for several other platforms including " _
                + "IBM PCs running DOS (1983), the Apple Macintosh (1984), the AT&T Unix PC (1985), " _
                + "Atari ST (1986), SCO UNIX, OS/2, and Microsoft Windows (1989). "
            Dim p2 As String _
                = "Microsoft Office Word instead of merely Microsoft Word. " _
                + "The 2010 version appears to be branded as Microsoft Word, " _
                + "once again. The current versions are Microsoft Word 2010 for Windows and 2008 for Mac."
            Dim paragraph As Paragraph = Nothing
            If section.Paragraphs.Count > 0 Then
                paragraph = section.Paragraphs(0)
            Else
                paragraph = section.AddParagraph()
            End If

            paragraph.AppendText(p1).CharacterFormat.FontSize = 14
            section.AddParagraph().AppendText(p2).CharacterFormat.FontSize = 14

            'Save doc file.
            document.SaveToFile("Sample.doc", FileFormat.Doc)

            'Launching the MS Word file.
            WordDocViewer("Sample.doc")
        End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
