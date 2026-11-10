Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Vba

Namespace ModifyMacro
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new Document object to work with Word files
			Dim doc As New Document()

			' Load an existing macro-enabled Word document from the specified relative path
			doc.LoadFromFile("..\..\..\..\..\..\Data\Macro.docm")

			' Retrieve the specific VBA module named "SampleModule2" from the document's VBA project
			Dim vbaModule As VbaModule = doc.VbaProject.Modules("SampleModule2")

			' Replace or set the source code of the VBA module with a new macro that indents a paragraph
			vbaModule.SourceCode = "" & ControlChars.CrLf & "Sub IndentParagraph()" & ControlChars.CrLf & "    Selection.ParagraphFormat.LeftIndent = InchesToPoints(2.5)" & ControlChars.CrLf & "End Sub"

			' Save the modified document as a new macro-enabled Word file (.docm)
			doc.SaveToFile("ModifyMacro.docm", Spire.Doc.FileFormat.Docm)

			' Close the document to release file handles
			doc.Close()

			' Dispose of the document object to free up unmanaged resources
			doc.Dispose()
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
