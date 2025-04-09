Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RecurseAllDocumentObject
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			' Create a new StringBuilder to store the content
			Dim builder As New StringBuilder()

			' Iterate through each Section in the document
			For Each section As Section In document.Sections
				' Get the index of the current section
				Dim SectionIndex As Integer = document.GetIndex(section)
				builder.AppendLine(String.Format("section index {0} has following ChildObjects", SectionIndex))

				' Iterate through each ChildObject in the section's body
				For Each obj As DocumentObject In section.Body.ChildObjects
					' Append the index and type of the ChildObject to the StringBuilder
					builder.AppendLine(String.Format("Index : {0}, ChildObject Type: {1}", section.Body.GetIndex(obj), obj.DocumentObjectType))
					
					' Check if the ChildObject is a Paragraph
					If obj.DocumentObjectType.Equals(DocumentObjectType.Paragraph) Then
						Dim paragraph As Paragraph = TryCast(obj, Paragraph)
						' Append the index and type of the Paragraph and its ChildObjects to the StringBuilder with indentation
						builder.AppendLine(String.Format(vbTab & "Paragraph index {0} has following ChildObjects", section.Body.GetIndex(paragraph)))
						For Each obj2 As DocumentObject In paragraph.ChildObjects
							builder.AppendLine(String.Format(vbTab & "Index : {0}, ChildObject Type: {1}", paragraph.GetIndex(obj2), obj2.DocumentObjectType))
						Next obj2
					End If
				Next obj
				builder.AppendLine(" ")
			Next section

			' Write the content of the StringBuilder to a text file
			File.WriteAllText("RecurseAllDocumentObject.txt", builder.ToString())

			' Release all resources used by the Document object
			document.Dispose()

			'Launching the Word file.
			TextViewer("RecurseAllDocumentObject.txt")


		End Sub

		Private Sub TextViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
