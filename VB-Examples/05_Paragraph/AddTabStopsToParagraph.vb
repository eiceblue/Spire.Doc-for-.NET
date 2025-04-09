Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AddTabStopsToParagraph
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object called document
			Dim document As New Document()

			'Add a new Section to the document and assign it to section
			Dim section As Section = document.AddSection()

			'Add a new Paragraph to section and assign it to paragraph1
			Dim paragraph1 As Paragraph = section.AddParagraph()

			'Add a Tab to the paragraph's format and assign it to tab
			Dim tab As Tab = paragraph1.Format.Tabs.AddTab(28)

			'Set the justification of the tab to Left
			tab.Justification = TabJustification.Left

			'Append text with a tab character and "Washing Machine" to paragraph1
			paragraph1.AppendText(vbTab & "Washing Machine")

			'Add another Tab to the paragraph's format and assign it to tab
			tab = paragraph1.Format.Tabs.AddTab(280)

			'Set the justification of the tab to Left
			tab.Justification = TabJustification.Left

			'Set the tab leader to Dotted
			tab.TabLeader = TabLeader.Dotted

			'Append text with a tab character and "$650" to paragraph1
			paragraph1.AppendText(vbTab & "$650")

			'Add a new Paragraph to section and assign it to paragraph2
			Dim paragraph2 As Paragraph = section.AddParagraph()

			'Add a Tab to the paragraph's format and assign it to tab
			tab = paragraph2.Format.Tabs.AddTab(28)

			'Set the justification of the tab to Left
			tab.Justification = TabJustification.Left

			'Append text with a tab character and "Refrigerator" to paragraph2
			paragraph2.AppendText(vbTab & "Refrigerator")

			'Add another Tab to the paragraph's format and assign it to tab
			tab = paragraph2.Format.Tabs.AddTab(280)

			'Set the justification of the tab to Left
			tab.Justification = TabJustification.Left

			'Set the tab leader to NoLeader
			tab.TabLeader = TabLeader.NoLeader

			'Append text with a tab character and "$800" to paragraph2
			paragraph2.AppendText(vbTab & "$800")

			'Specify the output file name
			Dim result As String = "Result-AddTabStopsToParagraph.docx"

			'Save the document as a Word document with the specified file format
			document.SaveToFile(result, FileFormat.Docx2013)

			'Dispose of the document object to release resources
			document.Dispose()

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
