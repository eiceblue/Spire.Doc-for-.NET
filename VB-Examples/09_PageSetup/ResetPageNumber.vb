Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ResetPageNumber
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create new instances of the Document class and load existing documents from file paths
			Dim document1 As New Document()
			document1.LoadFromFile("..\..\..\..\..\..\Data\ResetPageNumber1.docx")

			Dim document2 As New Document()
			document2.LoadFromFile("..\..\..\..\..\..\Data\ResetPageNumber2.docx")

			Dim document3 As New Document()
			document3.LoadFromFile("..\..\..\..\..\..\Data\ResetPageNumber3.docx")

			' Iterate through sections in document2 and document3, then add their clones to document1
			For Each sec As Section In document2.Sections
				document1.Sections.Add(sec.Clone())
			Next sec
			For Each sec As Section In document3.Sections
				document1.Sections.Add(sec.Clone())
			Next sec

			' Modify the footer fields in each section of document1
			For Each sec As Section In document1.Sections
				For Each obj As DocumentObject In sec.HeadersFooters.Footer.ChildObjects
					If obj.DocumentObjectType = DocumentObjectType.StructureDocumentTag Then
						Dim para As DocumentObject = obj.ChildObjects(0)
						For Each item As DocumentObject In para.ChildObjects
							If item.DocumentObjectType = DocumentObjectType.Field Then
								If (TryCast(item, Field)).Type = FieldType.FieldNumPages Then
									TryCast(item, Field).Type = FieldType.FieldSectionPages
								End If
							End If
						Next item
					End If
				Next obj
			Next sec

			' Reset page numbering for specific sections in document1
			document1.Sections(1).PageSetup.RestartPageNumbering = True
			document1.Sections(1).PageSetup.PageStartingNumber = 1
			document1.Sections(2).PageSetup.RestartPageNumbering = True
			document1.Sections(2).PageSetup.PageStartingNumber = 1

			' Specify the file name for the resulting document
			Dim result As String = "Result-ResetPageNumber.docx"

			' Save the modified document to a new file with the specified file format
			document1.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document objects
			document1.Dispose()
			document2.Dispose()
			document3.Dispose()

			'Launch the MS Word file.
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
