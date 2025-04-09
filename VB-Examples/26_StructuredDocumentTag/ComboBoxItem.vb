Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ComboBoxItem
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim input As String = "..\..\..\..\..\..\Data\ComboBox.docx"

			' Create a new document object
			Dim doc As New Document()

			' Load the document from the specified file path
			doc.LoadFromFile(input)

			' Iterate through each section in the document
			For Each section As Section In doc.Sections
				' Iterate through each document object in the section's body
				For Each bodyObj As DocumentObject In section.Body.ChildObjects
					' Check if the document object is a StructureDocumentTag
					If bodyObj.DocumentObjectType = DocumentObjectType.StructureDocumentTag Then
						' Check if the StructureDocumentTag is of type ComboBox
						If (TryCast(bodyObj, StructureDocumentTag)).SDTProperties.SDTType = SdtType.ComboBox Then
							' Access the ComboBox control properties
							Dim combo As SdtComboBox = TryCast((TryCast(bodyObj, StructureDocumentTag)).SDTProperties.ControlProperties, SdtComboBox)

							' Remove an item from the ComboBox
							combo.ListItems.RemoveAt(1)

							' Create a new SdtListItem and add it to the ComboBox
							Dim item As New SdtListItem("D", "D")
							combo.ListItems.Add(item)

							' Set the selected value of the ComboBox based on the item value "D"
							For Each sdtItem As SdtListItem In combo.ListItems
								If String.CompareOrdinal(sdtItem.Value, "D") = 0 Then
									combo.ListItems.SelectedValue = sdtItem
								End If
							Next sdtItem
						End If
					End If
				Next bodyObj
			Next section

			' Specify the output file name
			Dim output As String = "ComboBoxItem.docx"

			' Save the modified document to a file in Docx 2013 format
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
