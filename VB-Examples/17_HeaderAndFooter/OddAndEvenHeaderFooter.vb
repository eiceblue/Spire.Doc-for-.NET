Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace OddAndEvenHeaderFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\MultiplePages.docx"

			'Create a Word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first section
			Dim section As Section = doc.Sections(0)

			'Set the DifferentOddAndEvenPagesHeaderFooter property to ture
			section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = True

			'Add odd header
			Dim P3 As Paragraph = section.HeadersFooters.OddHeader.AddParagraph()

			'Append text
			Dim OH As TextRange = P3.AppendText("Odd Header")

			'Set the HorizontalAlignment for the paragraph
            P3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			'Set the font name and font size
			OH.CharacterFormat.FontName = "Arial"
			OH.CharacterFormat.FontSize = 10

			'Add even header
			Dim P4 As Paragraph = section.HeadersFooters.EvenHeader.AddParagraph()

			'Append text
			Dim EH As TextRange = P4.AppendText("Even Header from E-iceblue Using Spire.Doc")

			'Set the HorizontalAlignment for the paragraph
            P4.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			'Set the font name and font size
			EH.CharacterFormat.FontName = "Arial"
			EH.CharacterFormat.FontSize = 10

			'Add odd footer
			Dim P2 As Paragraph = section.HeadersFooters.OddFooter.AddParagraph()

			'Append text
			Dim [OF] As TextRange = P2.AppendText("Odd Footer")

			'Set the HorizontalAlignment for the paragraph
            P2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			'Set the font name and font size
			[OF].CharacterFormat.FontName = "Arial"
			[OF].CharacterFormat.FontSize = 10

			'Add even footer
			Dim P1 As Paragraph = section.HeadersFooters.EvenFooter.AddParagraph()

			'Append text
			Dim EF As TextRange = P1.AppendText("Even Footer from E-iceblue Using Spire.Doc")

			'Set the font name and font size
			EF.CharacterFormat.FontName = "Arial"
			EF.CharacterFormat.FontSize = 10

			'Set the HorizontalAlignment for the paragraph
            P1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			'Save the document
			Dim output As String = "OddAndEvenHeaderFooter.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
			doc.Dispose()
			
			Viewer(output)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
