Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ChangeTOCTabStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document
			Dim doc As New Document()

			' Load the document from a file
			doc.LoadFromFile("..\..\..\..\..\..\Data\Template_Toc.docx")

			' Create a custom Table of Contents (TOC) style
			Dim tocStyle As ParagraphStyle = TryCast(Style.CreateBuiltinStyle(BuiltinStyle.Toc1, doc), ParagraphStyle)
			tocStyle.CharacterFormat.FontName = "Aleo"
			tocStyle.CharacterFormat.FontSize = 15f
			tocStyle.CharacterFormat.TextColor = Color.CadetBlue
			doc.Styles.Add(tocStyle)

			' Iterate through all sections in the document
			For Each section As Section In doc.Sections
				' Iterate through all child objects in the body of each section
				For Each obj As DocumentObject In section.Body.ChildObjects
					' Check if the object is a StructureDocumentTag (e.g., TOC field)
					If TypeOf obj Is StructureDocumentTag Then
						Dim tag As StructureDocumentTag = TryCast(obj, StructureDocumentTag)

						' Iterate through all child objects within the StructureDocumentTag
						For Each [cObj] As DocumentObject In tag.ChildObjects
							' Check if the child object is a paragraph
							If TypeOf [cObj] Is Paragraph Then
								Dim para As Paragraph = TryCast([cObj], Paragraph)

								' Check if the paragraph has the style name "TOC1"
								If para.StyleName = "TOC1" Then
									' Apply the custom TOC style to the paragraph
									para.ApplyStyle(tocStyle.Name)
								End If
							End If
						Next [cObj]
					End If
				Next obj
			Next section

			' Specify the output file name
			Dim output As String = "ChangeTOCStyle_out.docx"

			' Save the modified document to a new file in DOCX format (version 2013)
			doc.SaveToFile(output, FileFormat.Docx2013)

			' Dispose of the document object
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
