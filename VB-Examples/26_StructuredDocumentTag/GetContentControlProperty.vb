Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace GetContentControlProperty
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            ' Specify the input file path
            Dim input As String = "..\..\..\..\..\..\Data\ContentControl.docx"

            ' Create a new document object
            Dim doc As New Document()

			' Load the document from the specified file path
			doc.LoadFromFile(input)

			' Get all the structure tags in the document
			Dim structureTags As StructureTags = GetAllTags(doc)

			' Initialize variables for storing tag properties
			Dim [alias] As String = Nothing
			Dim id As Decimal = 0
			Dim tag As String = Nothing
			Dim [property] As String = "Alias of contentControl" & vbTab & "ID          " & vbTab & "Tag             " & vbTab & "STDType        " & vbCr & "Content        " & vbCrLf
            Dim sdtType1 As String = Nothing
            Dim paragraph As Paragraph = Nothing
			Dim sdt As SdtType = SdtType.RichText
			Dim content As String = ""
			Dim textRange As TextRange = Nothing

			' Retrieve structure document tags and process their properties and content
			Dim tags As List(Of StructureDocumentTag) = structureTags.tags
			For i As Integer = 0 To tags.Count - 1
				[alias] = tags(i).SDTProperties.Alias
				id = tags(i).SDTProperties.Id
				tag = tags(i).SDTProperties.Tag
				sdt = tags(i).SDTProperties.SDTType
                sdtType1 = sdt.ToString()
                If sdt.Equals(SdtType.RichText) OrElse sdt.Equals(SdtType.Text) Then
                    If tags(i).ChildObjects.Count > 0 Then
                        For Each obj As DocumentObject In tags(i).ChildObjects
                            If TypeOf obj Is Paragraph Then
                                paragraph = TryCast(obj, Paragraph)
                                content &= paragraph.Text
                            End If
                        Next obj
                    End If
                End If
                [property] &= [alias] & "," & vbTab & id & "," & vbTab & tag & "," & vbTab & sdtType1 & "," & vbTab & content & vbCrLf
                content = ""
			Next i

			' Retrieve structure document tag inlines and process their properties and content
			Dim tagInlines As List(Of StructureDocumentTagInline) = structureTags.tagInlines
			For i As Integer = 0 To tagInlines.Count - 1
				[alias] = tagInlines(i).SDTProperties.Alias
				id = tagInlines(i).SDTProperties.Id
				tag = tagInlines(i).SDTProperties.Tag
				sdt = tagInlines(i).SDTProperties.SDTType
                sdtType1 = sdt.ToString()
                If sdt.Equals(SdtType.RichText) OrElse sdt.Equals(SdtType.Text) Then
                    If tagInlines(i).ChildObjects.Count > 0 Then
                        For Each obj As DocumentObject In tagInlines(i).ChildObjects
                            If TypeOf obj Is TextRange Then
                                textRange = TryCast(obj, TextRange)
                                content &= textRange.Text
                            End If
                        Next obj
                    End If
                End If
                [property] &= [alias] & "," & vbTab & id & "," & vbTab & tag & "," & vbTab & sdtType1 & "," & vbTab & content & vbCrLf
                content = ""
			Next i

			' Retrieve structure document tag rows and process their properties and content
			Dim rowTags As List(Of StructureDocumentTagRow) = structureTags.rowTags
			For i As Integer = 0 To rowTags.Count - 1
				[alias] = rowTags(i).SDTProperties.Alias
				id = rowTags(i).SDTProperties.Id
				tag = rowTags(i).SDTProperties.Tag
				sdt = rowTags(i).SDTProperties.SDTType
                sdtType1 = sdt.ToString()
                If sdt.Equals(SdtType.RichText) OrElse sdt.Equals(SdtType.Text) Then
                    If rowTags(i).ChildObjects.Count > 0 Then
                        For Each obj As DocumentObject In rowTags(i).ChildObjects
                            If TypeOf obj Is Paragraph Then
                                paragraph = TryCast(obj, Paragraph)
                                content &= paragraph.Text
                            End If
                        Next obj
                    End If
                End If
                [property] &= [alias] & "," & vbTab & id & "," & vbTab & tag & "," & vbTab & sdtType1 & "," & vbTab & content & vbCrLf
                content = ""
			Next i

			' Retrieve structure document tag cells and process their properties and content
			Dim cellTags As List(Of StructureDocumentTagCell) = structureTags.cellTags
			For i As Integer = 0 To cellTags.Count - 1
				[alias] = cellTags(i).SDTProperties.Alias
				id = cellTags(i).SDTProperties.Id
				tag = cellTags(i).SDTProperties.Tag
				sdt = cellTags(i).SDTProperties.SDTType
                sdtType1 = sdt.ToString()
                If sdt.Equals(SdtType.RichText) OrElse sdt.Equals(SdtType.Text) Then
                    If cellTags(i).ChildObjects.Count > 0 Then
                        For Each obj As DocumentObject In cellTags(i).ChildObjects
                            If TypeOf obj Is Paragraph Then
                                paragraph = TryCast(obj, Paragraph)
                                content &= paragraph.Text
                            End If
                        Next obj
                    End If
                End If
                [property] &= [alias] & "," & vbTab & id & "," & vbTab & tag & "," & vbTab & sdtType1 & "," & vbTab & content & vbCrLf
                content = ""
			Next i

			' Specify the output file name
			Dim output As String = "Property.txt"

			' Write the property string to the output file
			File.WriteAllText(output, [property].ToString())

			' Dispose the document object
			doc.Dispose()
			
			Viewer(output)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		'Get all StructureTags of the Word document
		Private Shared Function GetAllTags(ByVal document As Document) As StructureTags
			Dim structureTags As New StructureTags()
			For Each section As Section In document.Sections
				For Each obj As DocumentObject In section.Body.ChildObjects
					If obj.DocumentObjectType = DocumentObjectType.StructureDocumentTag Then
						structureTags.tags.Add(TryCast(obj, StructureDocumentTag))


					ElseIf obj.DocumentObjectType = DocumentObjectType.Paragraph Then
						For Each pobj As DocumentObject In (TryCast(obj, Paragraph)).ChildObjects
							If pobj.DocumentObjectType = DocumentObjectType.StructureDocumentTagInline Then
								structureTags.tagInlines.Add(TryCast(pobj, StructureDocumentTagInline))
							End If
						Next pobj
					ElseIf obj.DocumentObjectType = DocumentObjectType.Table Then
						For Each row As TableRow In (TryCast(obj, Table)).Rows
							If TypeOf row Is StructureDocumentTagRow Then
								structureTags.rowTags.Add(TryCast(row, StructureDocumentTagRow))
							End If
							For Each cell As TableCell In row.Cells
								If TypeOf cell Is StructureDocumentTagCell Then
									structureTags.cellTags.Add(TryCast(cell, StructureDocumentTagCell))
								End If
								For Each cellChild As DocumentObject In cell.ChildObjects
									If cellChild.DocumentObjectType = DocumentObjectType.StructureDocumentTag Then
										structureTags.tags.Add(TryCast(cellChild, StructureDocumentTag))
									ElseIf cellChild.DocumentObjectType = DocumentObjectType.Paragraph Then
										For Each pobj As DocumentObject In (TryCast(cellChild, Paragraph)).ChildObjects
											If pobj.DocumentObjectType = DocumentObjectType.StructureDocumentTagInline Then
												structureTags.tagInlines.Add(TryCast(pobj, StructureDocumentTagInline))
											End If
										Next pobj
									End If
								Next cellChild
							Next cell
						Next row
					End If
				Next obj
			Next section
			Return structureTags
		End Function
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
			Private m_tags As List(Of StructureDocumentTag)
			Public Property tags() As List(Of StructureDocumentTag)
				Get
					If m_tags Is Nothing Then
						m_tags = New List(Of StructureDocumentTag)()
					End If
					Return m_tags
				End Get
				Set(ByVal value As List(Of StructureDocumentTag))
					m_tags = value
				End Set
			End Property
			Private m_celltags As List(Of StructureDocumentTagCell)
			Public Property cellTags() As List(Of StructureDocumentTagCell)
				Get
					If m_celltags Is Nothing Then
						m_celltags = New List(Of StructureDocumentTagCell)()
					End If
					Return m_celltags
				End Get
				Set(ByVal value As List(Of StructureDocumentTagCell))
					m_celltags = value
				End Set
			End Property
			Private m_rowTags As List(Of StructureDocumentTagRow)
			Public Property rowTags() As List(Of StructureDocumentTagRow)
				Get
					If m_rowTags Is Nothing Then
						m_rowTags = New List(Of StructureDocumentTagRow)()
					End If
					Return m_rowTags
				End Get
				Set(ByVal value As List(Of StructureDocumentTagRow))
					m_rowTags = value
				End Set
			End Property
		End Class

	End Class
End Namespace
