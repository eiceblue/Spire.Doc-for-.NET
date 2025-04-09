Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.OMath

Namespace OfficeMathToOfficeMathMLCode
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim doc As New Document()
			' Load a Word document from a specific file path
			doc.LoadFromFile("..\..\..\..\..\..\Data\ToOfficeMathMLCode.docx")
			' Create a StringBuilder to store the MathML code
			Dim stringBuilder As New StringBuilder()
			' Iterate through sections in the document
			For Each section As Section In doc.Sections
				' Iterate through paragraphs in each section
				For Each par As Paragraph In section.Body.Paragraphs
					' Iterate through child objects in each paragraph
					For Each obj As DocumentObject In par.ChildObjects
						' Check if the object is an OfficeMath equation
						Dim omath As OfficeMath = TryCast(obj, OfficeMath)
						If omath Is Nothing Then
							Continue For
						End If
						' Convert OfficeMath equation to MathML code
						Dim mathml As String = omath.ToOfficeMathMLCode()
						' Append MathML code to the StringBuilder
						stringBuilder.Append(mathml)
						stringBuilder.Append(vbCrLf)
					Next obj
				Next par
			Next section
			' Write the MathML code to a text file
			File.WriteAllText("OfficeMathToOfficeMathMLCode.txt", stringBuilder.ToString())

			WordDocViewer("OfficeMathToOfficeMathMLCode.txt")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch e As Exception
				Debug.Write(e.StackTrace)
			End Try
		End Sub

	End Class
End Namespace
