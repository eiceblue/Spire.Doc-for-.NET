Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.IO
Imports Spire.Doc.Fields

Namespace ExtractTextFromTextBoxes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the Word document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\ExtractTextFromTextBoxes.docx")

			' Specify the output file name
			Dim result As String = "Result-ExtractTextFromTextBoxes.txt"

			' Check if the document contains any text boxes
			If document.TextBoxes.Count > 0 Then
				' Create a StreamWriter to write the extracted text to the output file
				Using sw As StreamWriter = File.CreateText(result)
					' Iterate through the sections in the document
					For Each section As Section In document.Sections
						' Iterate through the paragraphs in each section
						For Each p As Paragraph In section.Paragraphs
							' Iterate through the child objects of each paragraph
							For Each obj As DocumentObject In p.ChildObjects
								' Check if the child object is a text box
								If obj.DocumentObjectType = DocumentObjectType.TextBox Then
									' Cast the child object to a TextBox
									Dim textbox As Spire.Doc.Fields.TextBox = TryCast(obj, Spire.Doc.Fields.TextBox)

									' Iterate through the child objects of the text box
									For Each objt As DocumentObject In textbox.ChildObjects
										' Check if the child object is a paragraph
										If objt.DocumentObjectType = DocumentObjectType.Paragraph Then
											' Write the text of the paragraph to the output file
											sw.Write((TryCast(objt, Paragraph)).Text)
										End If

										' Check if the child object is a table
										If objt.DocumentObjectType = DocumentObjectType.Table Then
											' Cast the child object to a Table
											Dim table As Table = TryCast(objt, Table)

											' Extract text from the table and write it to the output file
											ExtractTextFromTables(table, sw)
										End If
									Next objt
								End If
							Next obj
						Next p
					Next section
				End Using
			End If

			' Dispose the Document object to free up resources
			document.Dispose()

			'Launch the result file.
			WordDocViewer(result)
		End Sub

		' Define a method to extract text from tables
		Private Shared Sub ExtractTextFromTables(ByVal table As Table, ByVal sw As StreamWriter)
			' Iterate through the rows of the table
			For i As Integer = 0 To table.Rows.Count - 1
				' Get the current row
				Dim row As TableRow = table.Rows(i)

				' Iterate through the cells of the row
				For j As Integer = 0 To row.Cells.Count - 1
					' Get the current cell
					Dim cell As TableCell = row.Cells(j)

					' Iterate through the paragraphs in the cell
					For Each paragraph As Paragraph In cell.Paragraphs
						' Write the text of the paragraph to the output file
						sw.Write(paragraph.Text)
					Next paragraph
				Next j
			Next i
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
