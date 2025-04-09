Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SplitDocByPageBreak
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object to hold the original document
			Dim original As New Document()

			' Load the original document from the specified file
			original.LoadFromFile("..\..\..\..\..\..\..\Data\SplitWordFileByPageBreak.docx")

			' Create a new Document object to store the split documents
			Dim newWord As New Document()

			' Add a new Section to the new document
			Dim section As Section = newWord.AddSection()

			' Clone default styles, themes, and compatibility settings from the original to the new document
			original.CloneDefaultStyleTo(newWord)
			original.CloneThemesTo(newWord)
			original.CloneCompatibilityTo(newWord)

			' Initialize the index for the split documents
			Dim index As Integer = 0

			' Iterate through each Section in the original document
			For Each sec As Section In original.Sections

				' Iterate through each DocumentObject in the Section's body
				For Each obj As DocumentObject In sec.Body.ChildObjects
					If TypeOf obj Is Paragraph Then
						' If the DocumentObject is a Paragraph, clone it along with the section properties to the new document
						Dim para As Paragraph = TryCast(obj, Paragraph)
						sec.CloneSectionPropertiesTo(section)
						section.Body.ChildObjects.Add(para.Clone())

						' Iterate through each child object in the paragraph
						For Each parobj As DocumentObject In para.ChildObjects
							If TypeOf parobj Is Break AndAlso (TryCast(parobj, Break)).BreakType = BreakType.PageBreak Then
								' If a page break is encountered, split the document and save the new segment as a separate file
								Dim i As Integer = para.ChildObjects.IndexOf(parobj)
								section.Body.LastParagraph.ChildObjects.RemoveAt(i)
								newWord.SaveToFile(String.Format("Result-SplitWordFileByPageBreak-{0}.docx", index), FileFormat.Docx)
								index += 1

								' Update the index and create a new Document and Section for the next segment
								newWord = New Document()
								section = newWord.AddSection()

								' Clone default styles, themes, and compatibility settings from the original to the new document
								original.CloneDefaultStyleTo(newWord)
								original.CloneThemesTo(newWord)
								original.CloneCompatibilityTo(newWord)

								' Clone the section properties and add the cloned paragraph to the section's body
								sec.CloneSectionPropertiesTo(section)
								section.Body.ChildObjects.Add(para.Clone())

								' Remove the paragraph content from the current section
								If section.Paragraphs(0).ChildObjects.Count = 0 Then
									section.Body.ChildObjects.RemoveAt(0)
								Else
									Do While i >= 0
										section.Paragraphs(0).ChildObjects.RemoveAt(i)
										i -= 1
									Loop
								End If
							End If
						Next parobj
					End If
					If TypeOf obj Is Table Then
						' If the DocumentObject is a Table, clone it to the new section's body
						section.Body.ChildObjects.Add(obj.Clone())
					End If
				Next obj
			Next sec

			' Save the last segment of the split document to a new file in Docx 2013 format
			Dim result As String = String.Format("Result-SplitWordFileByPageBreak-{0}.docx", index)
			newWord.SaveToFile(result, FileFormat.Docx2013)

			' Release all resources used by the Document objects
			original.Dispose()
			newWord.Dispose()

			'Launch the MS Word file.
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
