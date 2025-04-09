Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ASCIICharactersBulletStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim document As New Document()

			' Add a new section to the document
			Dim section As Section = document.AddSection()

			' Create and add list styles for bulleted lists
			Dim listStyle1 As New ListStyle(document, ListType.Bulleted)
			listStyle1.Name = "liststyle"
			listStyle1.Levels(0).BulletCharacter = ChrW(&H6E).ToString()
			listStyle1.Levels(0).CharacterFormat.FontName = "Wingdings"
			document.ListStyles.Add(listStyle1)

			Dim listStyle2 As New ListStyle(document, ListType.Bulleted)
			listStyle2.Name = "liststyle2"
			listStyle2.Levels(0).BulletCharacter = ChrW(&H75).ToString()
			listStyle2.Levels(0).CharacterFormat.FontName = "Wingdings"
			document.ListStyles.Add(listStyle2)

			Dim listStyle3 As New ListStyle(document, ListType.Bulleted)
			listStyle3.Name = "liststyle3"
			listStyle3.Levels(0).BulletCharacter = ChrW(&HB2).ToString()
			listStyle3.Levels(0).CharacterFormat.FontName = "Wingdings"
			document.ListStyles.Add(listStyle3)

			Dim listStyle4 As New ListStyle(document, ListType.Bulleted)
			listStyle4.Name = "liststyle4"
			listStyle4.Levels(0).BulletCharacter = ChrW(&HD8).ToString()
			listStyle4.Levels(0).CharacterFormat.FontName = "Wingdings"
			document.ListStyles.Add(listStyle4)

			' Add paragraphs with different bullet styles to the section
			Dim p1 As Paragraph = section.Body.AddParagraph()
			p1.AppendText("Spire.Doc for .NET")
			p1.ListFormat.ApplyStyle(listStyle1.Name)

			Dim p2 As Paragraph = section.Body.AddParagraph()
			p2.AppendText("Spire.Doc for .NET")
			p2.ListFormat.ApplyStyle(listStyle2.Name)

			Dim p3 As Paragraph = section.Body.AddParagraph()
			p3.AppendText("Spire.Doc for .NET")
			p3.ListFormat.ApplyStyle(listStyle3.Name)

			Dim p4 As Paragraph = section.Body.AddParagraph()
			p4.AppendText("Spire.Doc for .NET")
			p4.ListFormat.ApplyStyle(listStyle4.Name)

			' Save the document to a file
			Dim output As String = "ASCIICharactersBulletStyle_output.docx"
			document.SaveToFile(output, FileFormat.Docx)

			' Dispose of the document object
			document.Dispose()
			
			WordDocViewer(output)
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
