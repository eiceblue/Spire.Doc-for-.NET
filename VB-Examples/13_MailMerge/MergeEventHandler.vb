Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Interface
Imports Spire.Doc.Reporting

Namespace MergeEventHandler
	Partial Public Class Form1
		Inherits Form
		Private lastIndex As Integer = 0
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document instance
			Dim document As New Document()

			'Load the document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\Data\MergeEventHandler.doc")

			'Initialize the lastIndex variable
			lastIndex = 0

			'Create a list of CustomerRecord objects
			Dim customerRecords As New List(Of CustomerRecord)()

			'Add customer records to the list
			Dim c1 As New CustomerRecord()
			c1.ContactName = "Lucy"
			c1.Fax = "786-324-10"
			c1.Date = Date.Now
			customerRecords.Add(c1)


			Dim c2 As New CustomerRecord()
			c2.ContactName = "Lily"
			c2.Fax = "779-138-13"
			c2.Date = Date.Now
			customerRecords.Add(c2)

			Dim c3 As New CustomerRecord()
			c3.ContactName = "James"
			c3.Fax = "363-287-02"
			c3.Date = Date.Now
			customerRecords.Add(c3)

			'Subscribe to the MergeField event
			AddHandler document.MailMerge.MergeField, AddressOf MailMerge_MergeField

			'Execute the mail merge using the customerRecords list as the data source
			document.MailMerge.ExecuteGroup(New MailMergeDataTable("Customer", customerRecords))

			'Save doc file.
			document.SaveToFile("Sample.doc", FileFormat.Doc)

			'Dispose the document
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("Sample.doc")

		End Sub

		Private Sub MailMerge_MergeField(ByVal sender As Object, ByVal args As MergeFieldEventArgs)
			'Check if the current row index is greater than the lastIndex
			If args.RowIndex > lastIndex Then

				'Update the lastIndex with the current row index
				lastIndex = args.RowIndex

				'Add a page break before the current merge field
				AddPageBreakForMergeField(args.CurrentMergeField)
			End If
		End Sub
		
		Private Sub AddPageBreakForMergeField(ByVal mergeField As IMergeField)
			'Find position of needing to add page break
			Dim foundGroupStart As Boolean = False
			Dim paramgraph As Paragraph = TryCast(mergeField.PreviousSibling.Owner, Paragraph)
			Dim merageField As MergeField = Nothing

			'Find the group start merge field by traversing the previous sibling paragraphs
			Do While Not foundGroupStart
				paramgraph = TryCast(paramgraph.PreviousSibling, Paragraph)
				For i As Integer = 0 To paramgraph.Items.Count - 1
					merageField = TryCast(paramgraph.Items(i), MergeField)
					If (merageField IsNot Nothing) AndAlso (merageField.Prefix = "GroupStart") Then
						foundGroupStart = True
						Exit For
					End If
				Next i
			Loop

			'Append a page break to the paragraph
			paramgraph.AppendBreak(BreakType.PageBreak)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class

	Public Class CustomerRecord
		Private m_contactName As String
		Public Property ContactName() As String
			Get
				Return m_contactName
			End Get
			Set(ByVal value As String)
				m_contactName = value
			End Set
		End Property

		Private m_fax As String
		Public Property Fax() As String
			Get
				Return m_fax
			End Get
			Set(ByVal value As String)
				m_fax = value
			End Set
		End Property

		Private m_date As Date
		Public Property [Date]() As Date
			Get
				Return m_date
			End Get
			Set(ByVal value As Date)
				m_date = value
			End Set
		End Property
	End Class
End Namespace
