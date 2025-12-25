Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ModifyHyperlinkOfImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Input file path
			Dim input As String = "..\..\..\..\..\..\Data\ImageHyperlink.docx"

			'Output file path
			Dim output As String = "ModifyHyperlinkOfImage_output.docx"

			'Create word document
			Dim document As New Document()

			'Load a document
			document.LoadFromFile(input)

			' Iterate through each section in the document
			For Each section As Section In document.Sections
				' Iterate through each paragraph within the current section
				For Each paragraph As Paragraph In section.Paragraphs
					' Iterate through each child object within the current paragraph
					For Each documentObject As DocumentObject In paragraph.ChildObjects
						' Check if the current document object is a picture (DocPicture)
						If TypeOf documentObject Is DocPicture Then
							' Cast the document object to a DocPicture type
							Dim pic As DocPicture = TryCast(documentObject, DocPicture)

							' Check if the picture has a hyperlink associated with it
							If pic.HasHyperlink Then
								' Update the hyperlink of the picture to a new URL
								pic.HRef = "https://www.e-iceblue.com/Introduce/word-for-net-introduce.html"
							End If
						End If
					Next documentObject
				Next paragraph
			Next section


			' Save to file
			document.SaveToFile(output, FileFormat.Docx2019)

			'Dispose the document
			document.Dispose()

			'Launching the Word file.
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
