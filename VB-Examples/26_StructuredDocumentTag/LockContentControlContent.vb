Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace LockContentControlContent
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the HTML table string
			Dim htmlString As String = "<table style=""width: 100 % "">" & "<tr><th> Number </th><th> Name </th ><th>Age</th ></tr>" & "<tr><td> 1 </td><td> Smith </td><td> 50 </td></tr>" & "<tr> <td> 2 </td><td> Jackson </td><td> 94 </td> </tr>" & "</table>"

			' Create a new document
			Dim doc As New Document()

			' Add a section to the document
			Dim section As Section = doc.AddSection()

			' Add a paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append HTML content to the paragraph
			paragraph.AppendHTML(htmlString)

			' Create a StructureDocumentTag
			Dim sdt As New StructureDocumentTag(doc)

			' Add a new section to the document
			Dim section2 As Section = doc.AddSection()

			' Add the StructureDocumentTag to the section's body
			section2.Body.ChildObjects.Add(sdt)

			' Set the type of the StructureDocumentTag to RichText
			sdt.SDTProperties.SDTType = SdtType.RichText

			' Iterate through the child objects in the first section's body
			For Each obj As DocumentObject In section.Body.ChildObjects
				' Check if the object is a table
				If obj.DocumentObjectType = DocumentObjectType.Table Then
					' Clone and add the table to the StructureDocumentTag's content
					sdt.SDTContent.ChildObjects.Add(obj.Clone())
				End If
			Next obj

			' Lock the content editing settings of the StructureDocumentTag
			sdt.SDTProperties.LockSettings = LockSettingsType.ContentLocked

			' Remove the first section from the document
			doc.Sections.Remove(section)

			' Save the modified document to a file
			Dim result As String = "LockContentEditProperty_result.docx"
			doc.SaveToFile(result, Spire.Doc.FileFormat.Docx2013)

			' Dispose the document object
			doc.Dispose()
			
			'View the document
			FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
