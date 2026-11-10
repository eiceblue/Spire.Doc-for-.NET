Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace InserHtml
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object to represent an empty Word document.
			Dim doc As New Document()

			' Initialize a DocumentNavigator instance to simplify content insertion and navigation within the document.
			Dim navigator As New DocumentNavigator(doc)

			Dim html As String = "<body style=""font-family: Arial, sans-serif; background-color: #f4f4f9; color: #333; padding: 20px; line-height: 1.6;"">" & vbCrLf & "    <h1 style=""color: #2c3e50;"">Welcome to the Random English Page!</h1>" & vbCrLf & "    <p>This is a randomly generated HTML document containing English text.</p>" & vbCrLf & "    <p>The quick brown <span style=""background-color: #fffacd; padding: 4px 8px; border-radius: 4px;"">fox jumps</span> over the lazy dog. This sentence contains every letter of the English alphabet.</p>" & vbCrLf & "    <p>Here are a few fun facts:</p>" & vbCrLf & "    <ul>" & vbCrLf & "        <li>English is spoken by over 1.5 billion people worldwide.</li>" & vbCrLf & "        <li>The word ""set"" has the most definitions in the English language.</li>" & vbCrLf & "        <li>""Dreamt"" is the only English word that ends with ""mt"".</li>" & vbCrLf & "    </ul>" & vbCrLf & "    <p>Thank you for visiting! Have a wonderful day.</p>" & vbCrLf & "</body>"


			' Insert the HTML content (currently empty) into the document at the current cursor position.
			navigator.InsertHtml(html)

			' Save the resulting document to a file in DOCX format.
			doc.SaveToFile("InserHtml.docx", FileFormat.Docx)

			' Close the document to release internal resources.
			doc.Close()

			' Explicitly dispose of the document object to free memory immediately.
			doc.Dispose()

			'Launch the MS Word file.
			WordDocViewer("InserHtml.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
