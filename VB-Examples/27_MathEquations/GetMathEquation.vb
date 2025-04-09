Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.OMath

Namespace GetMathEquation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim doc As New Document()

			' Load a document from the specified file path
			doc.LoadFromFile("..\..\..\..\..\..\Data\GetMathEquation.docx")

			' Create a list to store the OfficeMath objects representing the math equations
			Dim mathEquations As New List(Of OfficeMath)()

			' Create a StringBuilder to build the MathML code as a string
			Dim stringBuilder As New StringBuilder()

			' Iterate through the sections in the document
			For Each section As Section In doc.Sections
				' Iterate through the paragraphs in the section
				For Each paragraph As Paragraph In section.Paragraphs
					' Iterate through the child objects in the paragraph
					For Each obj As DocumentObject In paragraph.ChildObjects
						' Check if the current object is an OfficeMath object
						If TypeOf obj Is OfficeMath Then
							' Append the MathML code of the OfficeMath object to the StringBuilder
							stringBuilder.AppendLine((TryCast(obj, OfficeMath)).ToMathMLCode())
							stringBuilder.AppendLine()

							' Add the OfficeMath object to the list
							mathEquations.Add(TryCast(obj, OfficeMath))
						End If
					Next obj
				Next paragraph
			Next section

			' Specify the output file path
			Dim output As String = "MathMLCode.txt"

			' Write the MathML code to the output file
			File.WriteAllText(output, stringBuilder.ToString())

			' Dispose the document object
			doc.Dispose()
			
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
