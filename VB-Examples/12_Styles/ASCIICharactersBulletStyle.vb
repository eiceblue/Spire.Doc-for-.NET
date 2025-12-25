Imports Spire.Doc
Imports Spire.Doc.Collections
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports System.IO
Imports System.Text

Namespace ASCIICharactersBulletStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			 'Create a new document
			Dim document As New Document()
			Dim section As Section = document.AddSection()

			'Create four list styles based on different ASCII characters
			Dim listStyle1 As ListStyle = document.Styles.Add(ListType.Bulleted, "liststyle")
			Dim Levels As ListLevelCollection = listStyle1.ListRef.Levels

			Levels(0).BulletCharacter = ChrW(&H006e).ToString()
			Levels(0).CharacterFormat.FontName = "Wingdings"

			'ListStyle listStyle2 = new ListStyle(document, ListType.Bulleted);
			'listStyle2.Name = "liststyle2";
			Dim listStyle2 As ListStyle = document.Styles.Add(ListType.Bulleted, "liststyle2")
			Dim Levels2 As ListLevelCollection = listStyle2.ListRef.Levels
			Levels2(0).BulletCharacter = ChrW(&H0075).ToString()
			Levels2(0).CharacterFormat.FontName = "Wingdings"

			Dim listStyle3 As ListStyle = document.Styles.Add(ListType.Bulleted, "liststyle3")
			Dim Levels3 As ListLevelCollection = listStyle3.ListRef.Levels
			Levels3(0).BulletCharacter = ChrW(&H00b2).ToString()
			Levels3(0).CharacterFormat.FontName = "Wingdings"

			Dim listStyle4 As ListStyle = document.Styles.Add(ListType.Bulleted, "liststyle4")
			Dim Levels4 As ListLevelCollection = listStyle4.ListRef.Levels
			Levels4(0).BulletCharacter = ChrW(&H00d8).ToString()
			Levels4(0).CharacterFormat.FontName = "Wingdings"

			'Add four paragraphs and apply list style separately
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

			'Save the document
			Dim output As String = "ASCIICharactersBulletStyle_output.docx"
			document.SaveToFile(output, FileFormat.Docx)

			'Dispose the Document
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
