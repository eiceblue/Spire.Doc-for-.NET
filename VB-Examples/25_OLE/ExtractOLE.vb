Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports System.IO

Namespace ExtractOLE
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim doc As New Document()

			' Load the document from a file
			doc.LoadFromFile("..\..\..\..\..\..\Data\OLEs.docx")

			' Iterate through each section in the document
			For Each sec As Section In doc.Sections
				' Iterate through each child object in the section's body
				For Each obj As DocumentObject In sec.Body.ChildObjects
					' Check if the object is a paragraph
					If TypeOf obj Is Paragraph Then
						' Cast the object to a paragraph
						Dim par As Paragraph = TryCast(obj, Paragraph)
						' Iterate through each child object in the paragraph
						For Each o As DocumentObject In par.ChildObjects
							' Check if the child object is an OLE object
							If o.DocumentObjectType = DocumentObjectType.OleObject Then
								' Cast the object to a DocOleObject
								Dim Ole As DocOleObject = TryCast(o, DocOleObject)
								' Get the type of the OLE object
								Dim s As String = Ole.ObjectType

								' Perform actions based on the OLE object type
								If s = "AcroExch.Document.DC" Then
									' Save the OLE object as a PDF file
									File.WriteAllBytes("Result.pdf", Ole.NativeData)
									' Open the PDF file with the default file viewer
									FileViewer("Result.pdf")
								ElseIf s = "Excel.Sheet.8" Then
									' Save the OLE object as an Excel file
									File.WriteAllBytes("ExcelResult.xls", Ole.NativeData)
									' Open the Excel file with the default file viewer
									FileViewer("ExcelResult.xls")
								ElseIf s = "PowerPoint.Show.12" Then
									' Save the OLE object as a PowerPoint file
									File.WriteAllBytes("PPTResult.pptx", Ole.NativeData)
									' Open the PowerPoint file with the default file viewer
									FileViewer("PPTResult.pptx")
								End If
							End If
						Next o
					End If
				Next obj
			Next sec

			' Dispose the document object
			doc.Dispose()
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
