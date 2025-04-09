Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Fields.OMath

Namespace ConvertEqToOfficeMath
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document
			Dim document As New Document()

			' Load the document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\EQ.docx")

			' Get the first paragraph of the first section in the document
			Dim paragraph As Paragraph = document.Sections(0).Paragraphs(0)

			' Iterate through the child objects of the paragraph
			Dim i As Integer = 0
			Do While i < paragraph.ChildObjects.Count
				' Get the current document object
				Dim documentObject As DocumentObject = paragraph.ChildObjects(i)

				' Check if the document object is a field of type Equation
				If TypeOf documentObject Is Field AndAlso (CType(documentObject, Field)).Type = FieldType.FieldEquation Then
					' Convert the field to an OfficeMath object
					Dim officeMath As OfficeMath = OfficeMath.FromEqField(CType(documentObject, Field))

					' If conversion is successful, replace the field with the OfficeMath object
					If officeMath IsNot Nothing Then
						paragraph.ChildObjects.Remove(documentObject)
						paragraph.ChildObjects.Insert(i, officeMath)
					End If
				End If
				i += 1
			Loop

			' Save the modified document to a new file
			document.SaveToFile("ConvertEqToOfficeMath.docx", FileFormat.Docx)

			' Dispose of the document object
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("ConvertEqToOfficeMath.docx")
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
