Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Reporting

Namespace AddHyperlinkForMailMergedImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim doc As New Document()
			' Load a Word document from a specific file path
			doc.LoadFromFile("..\..\..\..\..\..\Data\AddHyperlinkForImage.docx")
			' Define the field names and corresponding image file names
			Dim fieldNames = New String() { "MyImage" }
			Dim fieldValues = New String() { "..\..\..\..\..\..\Data\mailmerge_logo.png" }
			' Attach an event handler for the MergeImageField event
			AddHandler doc.MailMerge.MergeImageField, AddressOf MailMerge_MergeImageField
			' Execute the mail merge with the field names and values
			doc.MailMerge.Execute(fieldNames, fieldValues)
			' Save the modified document to a new file
			doc.SaveToFile("AddHyperlinkForImage.docx", FileFormat.Docx)


			WordDocViewer("AddHyperlinkForImage.docx")
		End Sub

		' Event handler for the MergeImageField event
		Private Sub MailMerge_MergeImageField(ByVal sender As Object, ByVal field As MergeImageFieldEventArgs)
			Dim filePath As String = field.ImageFileName ' FieldValue as string;
			If Not String.IsNullOrEmpty(filePath) Then
				field.Image = Image.FromFile(filePath)
				' Set the hyperlink for the merged image field
				field.ImageLink = "https://www.e-iceblue.com/"
			End If
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
