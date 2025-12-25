Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.OMath

Namespace AddMathEquation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Define an array of LaTeX math code strings
			Dim latexMathCode() As String = { "x^{2}+\sqrt{x^{2}+1}=2", "2\alpha - \sin y + x", "1 \over 2 + x", "(1 + \vert x-[a-b] \vert)", "\mbox{if $x=1$ or $x=2$}", "\begin{cases} 1 & \mbox{if $x>0$,} \\ 2 & \mbox{otherwise.} \end{cases}" }

			' Create a new document object
			Dim doc As New Document()

			' Load a document from the specified file path
			doc.LoadFromFile("..\..\..\..\..\..\Data\AddMathEquation.docx")

			' Get the first section in the document
			Dim section As Section = doc.Sections(0)

			' Declare variables for paragraph and OfficeMath objects
			Dim paragraph As Paragraph
			Dim officeMath As OfficeMath

			' Get the first table in the section
			Dim table1 As Table = TryCast(section.Tables(0), Table)

			' Create a list to store the OfficeMath objects representing the math equations
			Dim mathEquations As New List(Of OfficeMath)()

			' Iterate through the rows of the first table (excluding the header row)
			For i As Integer = 1 To 6
				' Get the first cell in the current row and add the LaTeX math code as text
				paragraph = table1.Rows(i).Cells(0).AddParagraph()
				paragraph.Text = latexMathCode(i - 1)

				' Get the second cell in the current row and create an OfficeMath object from the LaTeX math code
				paragraph = table1.Rows(i).Cells(1).AddParagraph()
				officeMath = New OfficeMath(doc)
				officeMath.FromLatexMathCode(latexMathCode(i - 1))
				paragraph.Items.Add(officeMath)

				' Add the OfficeMath object to the list
				mathEquations.Add(officeMath)
			Next i

			' Get the second table in the section
			Dim table2 As Table = TryCast(section.Tables(1), Table)

			' Iterate through the rows of the second table (excluding the header row)
			For i As Integer = 1 To 6
				' Get the first cell in the current row and add the MathML code of the corresponding OfficeMath object as text
				paragraph = table2.Rows(i).Cells(0).AddParagraph()
				paragraph.Text = mathEquations(i - 1).ToMathMLCode()

				' Get the second cell in the current row and create an OfficeMath object from the MathML code
				paragraph = table2.Rows(i).Cells(1).AddParagraph()
				officeMath = New OfficeMath(doc)
				officeMath.FromMathMLCode(mathEquations(i - 1).ToMathMLCode())
				paragraph.Items.Add(officeMath)
			Next i

			' Specify the output file path
			Dim result As String = "AddMathEquation_result.docx"

			' Save the modified document to the output file in DOCX format
			doc.SaveToFile(result, FileFormat.Docx)

			' Dispose the document object
			doc.Dispose()
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
