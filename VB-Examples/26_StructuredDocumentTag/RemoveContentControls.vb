Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace RemoveContentControls
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim doc As New Document()

			' Load a document file from a specified path
			doc.LoadFromFile("..\..\..\..\..\..\Data\RemoveContentControls.docx")

			' Iterate through the sections in the document
			For s As Integer = 0 To doc.Sections.Count - 1
				' Get the current section
				Dim section As Section = doc.Sections(s)

				' Iterate through the child objects in the section's body
				Dim i As Integer = 0
				Do While i < section.Body.ChildObjects.Count
					' Check if the child object is a paragraph
					If TypeOf section.Body.ChildObjects(i) Is Paragraph Then
						' Get the paragraph object
						Dim para As Paragraph = TryCast(section.Body.ChildObjects(i), Paragraph)

						' Iterate through the child objects in the paragraph
						Dim j As Integer = 0
						Do While j < para.ChildObjects.Count
							' Check if the child object is a StructureDocumentTagInline
							If TypeOf para.ChildObjects(j) Is StructureDocumentTagInline Then
								' Get the StructureDocumentTagInline object
								Dim sdt As StructureDocumentTagInline = TryCast(para.ChildObjects(j), StructureDocumentTagInline)

								' Remove the StructureDocumentTagInline from the paragraph
								para.ChildObjects.Remove(sdt)

								' Decrement the index to account for the removed object
								j -= 1
							End If
							j += 1
						Loop
					End If

					' Check if the child object is a StructureDocumentTag
					If TypeOf section.Body.ChildObjects(i) Is StructureDocumentTag Then
						' Get the StructureDocumentTag object
						Dim sdt As StructureDocumentTag = TryCast(section.Body.ChildObjects(i), StructureDocumentTag)

						' Remove the StructureDocumentTag from the section's body
						section.Body.ChildObjects.Remove(sdt)

						' Decrement the index to account for the removed object
						i -= 1
					End If
					i += 1
				Loop
			Next s

			' Save the modified document to a new file
			Dim output As String = "RemoveContentControls_out.docx"
			doc.SaveToFile(output, FileFormat.Docx2013)

			' Dispose the document object
			doc.Dispose()

			'Launch the file
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
