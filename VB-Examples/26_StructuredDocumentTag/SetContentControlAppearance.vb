Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetContentControlAppearance
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim input As String = "..\..\..\..\..\..\Data\ContentControl.docx"

			' Create a new document object
			Dim doc As New Document()

			' Load a document from the specified input file
			doc.LoadFromFile(input)

			' Iterate through the sections in the document
			For Each section As Section In doc.Sections
				' Iterate through the child objects in the section's body
				For Each docObj As DocumentObject In section.Body.ChildObjects
					' Check if the current object is a StructureDocumentTag
					If TypeOf docObj Is StructureDocumentTag Then
						' Get the StructureDocumentTag object and its SDTProperties
						Dim stdTag As StructureDocumentTag = CType(docObj, StructureDocumentTag)
						Dim sDTProperties As SDTProperties = stdTag.SDTProperties

						' Set the appearance of the StructureDocumentTag based on its SDTType
						Select Case sDTProperties.SDTType
							Case SdtType.Text
								sDTProperties.Appearance = SdtAppearance.BoundingBox
							Case SdtType.RichText
								sDTProperties.Appearance = SdtAppearance.Hidden
							Case SdtType.Picture
								sDTProperties.Appearance = SdtAppearance.Tags
							Case SdtType.CheckBox
								sDTProperties.Appearance = SdtAppearance.Default
						End Select
					End If
				Next docObj
			Next section

			' Specify the output file path
			Dim output As String = "SetContentControlAppearance.docx"

			' Save the modified document to the output file
			doc.SaveToFile(output, FileFormat.Docx2013)

			' Dispose the document object
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
