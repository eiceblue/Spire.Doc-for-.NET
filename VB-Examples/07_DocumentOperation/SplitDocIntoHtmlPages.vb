Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO
Imports Spire.Doc.Fields
Namespace SplitDocIntoHtmlPages
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Define the input file path
			Dim input As String = "..\..\..\..\..\..\..\Data\SplitDocIntoHtmlPages.doc"

			' Define the output directory path and create the directory if it doesn't exist
			Dim outDir As String = Path.Combine("output")
			Directory.CreateDirectory(outDir)

			' Call the method to split the document into multiple HTML pages
			SplitDocIntoMultipleHtml(input, outDir)
			End Sub

			' Method to split the input document into multiple HTML pages
			Private Sub SplitDocIntoMultipleHtml(ByVal input As String, ByVal outDirectory As String)
				' Create a new Document object and load the input document
				Dim document As New Document()
				document.LoadFromFile(input)

				' Initialize variables for sub-document, first section flag, and index
				Dim subDoc As Document = Nothing
				Dim first As Boolean = True
				Dim index As Integer = 0

				' Iterate through each Section in the document
				For Each sec As Section In document.Sections
					' Iterate through each DocumentObject in the Section's body
					For Each element As DocumentObject In sec.Body.ChildObjects
						' Check if the current element indicates the start of a new sub-document
						If IsInNextDocument(element) Then
							' If not the first sub-document, save the previous sub-document as an HTML file
							If Not first Then
								subDoc.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal
								subDoc.HtmlExportOptions.ImageEmbedded = True
								subDoc.SaveToFile(Path.Combine(outDirectory, String.Format("out-{0}.html", index)), FileFormat.Html)
								index += 1
								subDoc = Nothing
							End If
							first = False
						End If
						' Create a new sub-document if necessary and add the element to its body
						If subDoc Is Nothing Then
							subDoc = New Document()
							subDoc.AddSection()
						End If
						subDoc.Sections(0).Body.ChildObjects.Add(element.Clone())
					Next element
				Next sec

				' Save the final sub-document as an HTML file if it exists
				If subDoc IsNot Nothing Then
					subDoc.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal
					subDoc.HtmlExportOptions.ImageEmbedded = True
					subDoc.SaveToFile(Path.Combine(outDirectory, String.Format("out-{0}.html", index)), FileFormat.Html)
					index += 1
				End If
			End Sub

			' Function to determine if the element marks the start of a new sub-document
			Private Function IsInNextDocument(ByVal element As DocumentObject) As Boolean
				If TypeOf element Is Paragraph Then
					Dim p As Paragraph = TryCast(element, Paragraph)
					If p.StyleName = "Heading1" Then
						Return True
					End If
				End If
				Return False
			End Function
	End Class
End Namespace
