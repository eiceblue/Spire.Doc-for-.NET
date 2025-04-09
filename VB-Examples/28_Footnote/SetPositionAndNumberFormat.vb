Imports Spire.Doc

Namespace SetPositionAndNumberFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim input As String = "..\..\..\..\..\..\Data\Footnote.docx"

			' Create a new instance of Document
			Dim doc As New Document()

			' Load the Word document from the specified input file
			doc.LoadFromFile(input)

			' Get the first section of the document
			Dim sec As Section = doc.Sections(0)

			' Set the footnote options for the section
			sec.FootnoteOptions.NumberFormat = FootnoteNumberFormat.UpperCaseLetter
			sec.FootnoteOptions.RestartRule = FootnoteRestartRule.RestartPage
			sec.FootnoteOptions.Position = FootnotePosition.PrintAsEndOfSection

			' Specify the output file path
			Dim output As String = "SetPositionAndNumberFormat.docx"

			' Save the modified document to the specified output file
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose of the document object when finished using it
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
