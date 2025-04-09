Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ModifySDTColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim doc As New Document()

			' Load a document from the specified file path
			doc.LoadFromFile("..\..\..\..\..\..\Data\ModifySTDColor.docx")

			' Iterate through the sections in the document
			For s As Integer = 0 To doc.Sections.Count - 1
				' Get the current section
				Dim section As Section = doc.Sections(s)

				' Iterate through the child objects in the section's body
				For i As Integer = 0 To section.Body.ChildObjects.Count - 1
					' Check if the child object is a Paragraph
					If TypeOf section.Body.ChildObjects(i) Is Paragraph Then
						' Get the paragraph object
						Dim para As Paragraph = TryCast(section.Body.ChildObjects(i), Paragraph)

						' Iterate through the child objects in the paragraph
						For j As Integer = 0 To para.ChildObjects.Count - 1
							' Check if the child object is a StructureDocumentTagInline
							If TypeOf para.ChildObjects(j) Is StructureDocumentTagInline Then
								' Get the StructureDocumentTagInline object
								Dim sdt As StructureDocumentTagInline = TryCast(para.ChildObjects(j), StructureDocumentTagInline)

								' Get the SDTProperties of the StructureDocumentTagInline
								Dim sDTProperties As SDTProperties = sdt.SDTProperties

								' Set the color of the SDTProperties based on the SDTType
								Select Case sDTProperties.SDTType
									Case SdtType.RichText
										sDTProperties.Color = Color.Orange
									Case SdtType.Text
										sDTProperties.Color = Color.Green
								End Select
							End If
						Next j
					End If

					' Check if the child object is a StructureDocumentTag
					If TypeOf section.Body.ChildObjects(i) Is StructureDocumentTag Then
						' Get the StructureDocumentTag object
						Dim sdt As StructureDocumentTag = TryCast(section.Body.ChildObjects(i), StructureDocumentTag)

						' Get the SDTProperties of the StructureDocumentTag
						Dim sDTProperties As SDTProperties = sdt.SDTProperties

						' Set the color of the SDTProperties based on the SDTType
						Select Case sDTProperties.SDTType
							Case SdtType.RichText
								sDTProperties.Color = Color.Orange
							Case SdtType.Text
								sDTProperties.Color = Color.Green
						End Select
					End If
				Next i
			Next s

			' Specify the output file path
			Dim output As String = "ModifySTDColor_out.docx"

			' Save the modified document to the output file in DOCX format
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
