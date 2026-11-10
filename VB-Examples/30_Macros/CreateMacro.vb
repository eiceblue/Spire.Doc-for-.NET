Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Vba

Namespace CreateMacro
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Word document instance
			Dim document As New Document()

			' Add a new section to the document
			Dim section As Section = document.AddSection()

			' Add a new paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append text content to the paragraph
			paragraph.AppendText("There is a VBA Macro. ")

			' Create a new VBA project to hold macros
			Dim vbaProject As New VbaProject()

			' Set the name of the VBA project
			vbaProject.Name = "SampleVBAMacro"

			' Assign the VBA project to the document
			document.VbaProject = vbaProject

			' Add a standard VBA module named "SampleModule1" to the project
			Dim vbaModule1 As VbaModule = document.VbaProject.Modules.Add("SampleModule1", Spire.Doc.Vba.VbaModuleType.StdModule)

			' Define the VBA source code for the first module (displays document info and writes "Hello World!")
			vbaModule1.SourceCode = "" & ControlChars.CrLf & " Sub DocumnetInfo()" & ControlChars.CrLf & "     MsgBox ""create time: "" &Now()" & ControlChars.CrLf & "     MsgBox ""Pages:"" & ActiveDocument.Range.ComputeStatistics(wdStatisticPages)" & ControlChars.CrLf & " End Sub" & ControlChars.CrLf & ControlChars.CrLf & " Sub WriteHello()" & ControlChars.CrLf & "     Selection.TypeText Text:=""Hello World!""" & ControlChars.CrLf & " End Sub"

			' Add a second standard VBA module named "SampleModule2" to the project
			Dim vbaModule2 As VbaModule = document.VbaProject.Modules.Add("SampleModule2", Spire.Doc.Vba.VbaModuleType.StdModule)

			' Define the VBA source code for the second module (inserts current date and indents a paragraph)
			vbaModule2.SourceCode = "" & ControlChars.CrLf & " Sub InsertCurrentDate()" & ControlChars.CrLf & "     Selection.TypeText Text:=Format(Now(),""yyyy-mm-dd hh:mm:ss"")" & ControlChars.CrLf & " End Sub" & ControlChars.CrLf & ControlChars.CrLf & " Sub IndentParagraph()" & ControlChars.CrLf & "     Selection.ParagraphFormat.LeftIndent = InchesToPoints(0.5)" & ControlChars.CrLf & " End Sub"

			' Save the document as a macro-enabled Word file (.docm)
			document.SaveToFile("CreateMacro.docm", Spire.Doc.FileFormat.Docm)

			' Close the document and release resources
			document.Close()

			' Dispose of the document object and release any associated resources
			document.Dispose()

			'Launching the Word file.
			WordDocViewer("CreateMacro.docm")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
