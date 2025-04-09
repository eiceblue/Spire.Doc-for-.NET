Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO

Namespace CheckFileFormat
    Partial Public Class Form1
        Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            ' Load the input file path into a variable
			Dim input As String = "..\..\..\..\..\..\Data\Template.docx"

			' Create a new Document object
			Dim doc As New Document()

			' Load the document from the input file
			doc.LoadFromFile(input)

			' Get the detected file format of the document
			Dim ff As FileFormat = doc.DetectedFormatType

			' Initialize a variable to store the file format description
			Dim file As String = "The file format is "

			' Check the detected file format and assign the corresponding description
			Select Case ff
				Case FileFormat.Doc
					file &= "Microsoft Word 97-2003 document."
				Case FileFormat.Dot
					file &= "Microsoft Word 97-2003 template."
				Case FileFormat.Docx
					file &= "Office Open XML WordprocessingML Macro-Free Document."
				Case FileFormat.Docm
					file &= "Office Open XML WordprocessingML Macro-Enabled Document."
				Case FileFormat.Dotx
					file &= "Office Open XML WordprocessingML Macro-Free Template."
				Case FileFormat.Dotm
					file &= "Office Open XML WordprocessingML Macro-Enabled Template."
				Case FileFormat.Rtf
					file &= "RTF format."
				Case FileFormat.WordML
					file &= "Microsoft Word 2003 WordprocessingML format."
				Case FileFormat.Html
					file &= "HTML format."
				Case FileFormat.WordXml
					file &= "Microsoft Word xml format for word 2007-2013."
				Case FileFormat.Odt
					file &= "OpenDocument Text."
				Case FileFormat.Ott
					file &= "OpenDocument Text Template."
				Case FileFormat.DocPre97
					file &= "Microsoft Word 6 or Word 95 format."
				Case Else
					file &= "Unknown format."
			End Select

            MessageBox.Show(file)
        End Sub
    End Class
End Namespace
