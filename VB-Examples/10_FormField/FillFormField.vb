Imports System.ComponentModel
Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace FillFormField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            'open form
            Dim document As New Document("..\..\..\..\..\..\Data\UserForm.doc")

            'load data
            Using stream As Stream = File.OpenRead("..\..\..\..\..\..\Data\User.xml")
                Dim xpathDoc As New XPathDocument(stream)
                Dim user As XPathNavigator = xpathDoc.CreateNavigator().SelectSingleNode("/user")

                'fill data
                For Each field As FormField In document.Sections(0).Body.FormFields
                    Dim path As [String] = [String].Format("{0}/text()", field.Name)
                    Dim propertyNode As XPathNavigator = user.SelectSingleNode(path)
                    If propertyNode IsNot Nothing Then
                        Select Case field.Type
                            Case FieldType.FieldFormTextInput
                                field.Text = propertyNode.Value
                                Exit Select

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
                                Next
                                Exit Select

                            Case FieldType.FieldFormCheckBox
                                If Convert.ToBoolean(propertyNode.Value) Then
                                    Dim checkBox As CheckBoxFormField = TryCast(field, CheckBoxFormField)
                                    checkBox.Checked = True
                                End If
                                Exit Select
                        End Select
                    End If
                Next
            End Using

			'Save doc file.
            document.SaveToFile("Sample.doc", FileFormat.Doc)

			'Launching the MS Word file.
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
