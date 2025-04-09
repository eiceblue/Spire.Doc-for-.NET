Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetCaptionWithChapterNumber
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim document As New Document()

			' Load the Word document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\Data\SetCaptionWithChapterNumber.docx")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Specify the base name for the captions
			Dim name As String = "Caption "

			' Iterate through paragraphs in the body of the section
			For i As Integer = 0 To section.Body.Paragraphs.Count - 1
				' Iterate through child objects within each paragraph
				For j As Integer = 0 To section.Body.Paragraphs(i).ChildObjects.Count - 1
					' Check if the child object is a picture
					If TypeOf section.Body.Paragraphs(i).ChildObjects(j) Is DocPicture Then
						' Convert the child object to a DocPicture
						Dim pic1 As DocPicture = TryCast(section.Body.Paragraphs(i).ChildObjects(j), DocPicture)

						' Get the owner paragraph's owner, which should be the Body
						Dim body As Body = TryCast(pic1.OwnerParagraph.Owner, Body)

						If body IsNot Nothing Then
							' Find the index of the owner paragraph within the Body
							Dim imageIndex As Integer = body.ChildObjects.IndexOf(pic1.OwnerParagraph)

							' Create a new paragraph
							Dim para As New Paragraph(document)

							' Append the caption name
							para.AppendText(name)

							' Append a field for referencing the chapter number using a style reference
							Dim field1 As Field = para.AppendField("test", FieldType.FieldStyleRef)
							field1.Code = " STYLEREF 1 \s "

							' Append a separator text
							para.AppendText(" - ")

							' Append a sequence field for the caption number
							Dim field2 As SequenceField = CType(para.AppendField(name, FieldType.FieldSequence), SequenceField)
							field2.CaptionName = name
							field2.NumberFormat = CaptionNumberingFormat.Number

							' Insert the new paragraph after the owner paragraph
							body.Paragraphs.Insert(imageIndex + 1, para)
						End If
					End If
				Next j
			Next i

			' Enable field updating in the document
			document.IsUpdateFields = True

			' Specify the output file name and format (Docx)
			Dim output As String = "SetCaptionWithChapterNumber.docx"
			document.SaveToFile(output, FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launching the file
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
