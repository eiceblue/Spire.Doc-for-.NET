Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports System.Xml.XPath

Namespace FillFormField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a file
			Dim document As New Document("..\..\..\..\..\..\Data\FillFormField.doc")

			' Open the XML file containing user data
			Using stream As Stream = File.OpenRead("..\..\..\..\..\..\Data\User.xml")
				' Create an XPathDocument from the XML stream
				Dim xpathDoc As New XPathDocument(stream)

				' Select the "user" node from the XML document
				Dim user As XPathNavigator = xpathDoc.CreateNavigator().SelectSingleNode("/user")

				' Iterate through each form field in the document's first section
				For Each field As FormField In document.Sections(0).Body.FormFields
					' Get the XPath to retrieve the value for the current form field
					Dim path As String = String.Format("{0}/text()", field.Name)

					' Select the corresponding node from the XML document
					Dim propertyNode As XPathNavigator = user.SelectSingleNode(path)

					' If the node exists, set the value of the form field based on its type
					If propertyNode IsNot Nothing Then
						Select Case field.Type
							' Text input field
							Case FieldType.FieldFormTextInput
								field.Text = propertyNode.Value

							' Dropdown field
							Case FieldType.FieldFormDropDown
								Dim combox As DropDownFormField = TryCast(field, DropDownFormField)
								For i As Integer = 0 To combox.DropDownItems.Count - 1
									If combox.DropDownItems(i).Text = propertyNode.Value Then
										combox.DropDownSelectedIndex = i
										Exit For
									End If
									If field.Name = "country" AndAlso combox.DropDownItems(i).Text = "Others" Then
										combox.DropDownSelectedIndex = i
									End If
								Next i

							' Checkbox field
							Case FieldType.FieldFormCheckBox
								If Convert.ToBoolean(propertyNode.Value) Then
									Dim checkBox As CheckBoxFormField = TryCast(field, CheckBoxFormField)
									checkBox.Checked = True
								End If
						End Select
					End If
				Next field
			End Using

			' Save the modified document to a file
			document.SaveToFile("Sample.doc", FileFormat.Doc)

			' Dispose the document object
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("Sample.doc")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
