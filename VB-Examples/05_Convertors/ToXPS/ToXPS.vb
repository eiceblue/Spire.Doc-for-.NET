Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ToXPS
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

	Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
	'Create word document
	Dim document As New Document()

	Dim section As Section = document.AddSection()
	section.PageSetup.PageSize = PageSize.A4
	section.PageSetup.Margins.Top = 72F
	section.PageSetup.Margins.Bottom = 72F
	section.PageSetup.Margins.Left = 89.85F
	section.PageSetup.Margins.Right = 89.85F

	Dim paragraph As Paragraph = section.AddParagraph()
	paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left
            paragraph.AppendPicture(My.Resources.Word)

            Dim p1 As String = "Microsoft Word is a word processor designed by Microsoft. " & "It was first released in 1983 under the name Multi-Tool Word for Xenix systems. " & "Subsequent versions were later written for several other platforms including " & "IBM PCs running DOS (1983), the Apple Macintosh (1984), the AT&T Unix PC (1985), " & "Atari ST (1986), SCO UNIX, OS/2, and Microsoft Windows (1989). "
            Dim p2 As String = "Microsoft Office Word instead of merely Microsoft Word. " & "The 2010 version appears to be branded as Microsoft Word, " & "once again. The current versions are Microsoft Word 2010 for Windows and 2008 for Mac."
	section.AddParagraph().AppendText(p1).CharacterFormat.FontSize = 14
	section.AddParagraph().AppendText(p2).CharacterFormat.FontSize = 14

            'Save doc file to xps file.
	document.SaveToFile("Sample.xps", FileFormat.XPS)

	'Launching the xps file.
	FileViewer("Sample.xps")
End Sub

Private Sub FileViewer(ByVal fileName As String)
	Try
		System.Diagnostics.Process.Start(fileName)
	Catch
	End Try
End Sub
	End Class
End Namespace
