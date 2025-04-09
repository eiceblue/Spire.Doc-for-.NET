Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace UpdateCheckBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim document As New Document()

			' Load a document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\CheckBoxContentControl.docx")

			' Get all the StructureTags from the document
			Dim structureTags As StructureTags = GetAllTags(document)

			' Get the list of StructureDocumentTagInline objects from the StructureTags
			Dim tagInlines As List(Of StructureDocumentTagInline) = structureTags.tagInlines

			' Iterate through the list of StructureDocumentTagInline objects
			For i As Integer = 0 To tagInlines.Count - 1
				' Get the SDTType of the current StructureDocumentTagInline
				Dim type As String = tagInlines(i).SDTProperties.SDTType.ToString()

				' Check if the SDTType is "CheckBox"
				If type = "CheckBox" Then
					' Get the SdtCheckBox from the ControlProperties of the StructureDocumentTagInline
					Dim scb As SdtCheckBox = TryCast(tagInlines(i).SDTProperties.ControlProperties, SdtCheckBox)

					' Toggle the Checked property of the SdtCheckBox
					If scb.Checked Then
						scb.Checked = False
					Else
						scb.Checked = True
					End If
				End If
			Next i

			' Save the modified document to "Output.docx" in DOCX format
			document.SaveToFile("Output.docx", FileFormat.Docx)

			' Dispose the document object
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")

		End Sub

		' Define a method named "GetAllTags" that takes a Document object as input and returns a StructureTags object
		Private Shared Function GetAllTags(ByVal document As Document) As StructureTags
			' Create a new StructureTags object to store the StructureDocumentTagInline objects
			Dim structureTags As New StructureTags()

			' Iterate through the sections in the document
			For Each section As Section In document.Sections
				' Iterate through the child objects in the section's body
				For Each obj As DocumentObject In section.Body.ChildObjects
					' Check if the current object is a Paragraph
					If obj.DocumentObjectType = DocumentObjectType.Paragraph Then
						' Iterate through the child objects in the paragraph
						For Each pobj As DocumentObject In (TryCast(obj, Paragraph)).ChildObjects
							' Check if the current object is a StructureDocumentTagInline
							If pobj.DocumentObjectType = DocumentObjectType.StructureDocumentTagInline Then
								' Add the StructureDocumentTagInline to the tagInlines list in the StructureTags object
								structureTags.tagInlines.Add(TryCast(pobj, StructureDocumentTagInline))
							End If
						Next pobj
					End If
				Next obj
			Next section

			' Return the StructureTags object containing the collected StructureDocumentTagInline objects
			Return structureTags
		End Function

		' Define a public class named "StructureTags"
		Public Class StructureTags
			Private m_tagInlines As List(Of StructureDocumentTagInline)
			Public Property tagInlines() As List(Of StructureDocumentTagInline)
				Get
					If m_tagInlines Is Nothing Then
						m_tagInlines = New List(Of StructureDocumentTagInline)()
					End If
					Return m_tagInlines
				End Get
				Set(ByVal value As List(Of StructureDocumentTagInline))
					m_tagInlines = value
				End Set
			End Property
		End Class
		
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
