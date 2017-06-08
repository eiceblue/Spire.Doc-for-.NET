Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Interface
Imports Spire.Doc.Reporting

Namespace MergeEventHandler
	Partial Public Class Form1
        Inherits Form

        Private lastIndex As Integer = 0
        Private WithEvents mailMerge As MailMerge

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
            Dim document As New Document()
            document.LoadFromFile("..\..\..\..\..\..\Data\Fax2.doc")
            lastIndex = 0

            Dim customerRecords As New List(Of CustomerRecord)()
            Dim c1 As New CustomerRecord()
            c1.ContactName = "Lucy"
            c1.Fax = "786-324-10"
            c1.[Date] = DateTime.Now
            customerRecords.Add(c1)

            Dim c2 As New CustomerRecord()
            c2.ContactName = "Lily"
            c2.Fax = "779-138-13"
            c2.[Date] = DateTime.Now
            customerRecords.Add(c2)

            Dim c3 As New CustomerRecord()
            c3.ContactName = "James"
            c3.Fax = "363-287-02"
            c3.[Date] = DateTime.Now
            customerRecords.Add(c3)

            'Execute mailmerge
            mailMerge = document.MailMerge
            document.MailMerge.ExecuteGroup(New MailMergeDataTable("Customer", customerRecords))

            'Save doc file.
            document.SaveToFile("Sample.doc", FileFormat.Doc)

            'Launching the MS Word file.
            WordDocViewer("Sample.doc")
        End Sub

        Private Sub MailMerge_MergeField(ByVal sender As Object, ByVal args As MergeFieldEventArgs) Handles mailMerge.MergeField
            'Next row
            If args.RowIndex > lastIndex Then
                lastIndex = args.RowIndex
                AddPageBreakForMergeField(args.CurrentMergeField)
            End If
        End Sub

        Private Sub AddPageBreakForMergeField(ByVal mergeField As IMergeField)
            'Find position of needing to add page break
            Dim foundGroupStart As Boolean = False
            Dim paramgraph As Paragraph = TryCast(mergeField.PreviousSibling.Owner, Paragraph)
            Dim merageField As MergeField = Nothing
            While Not foundGroupStart
                paramgraph = TryCast(paramgraph.PreviousSibling, Paragraph)
                For i As Integer = 0 To paramgraph.Items.Count - 1
                    merageField = TryCast(paramgraph.Items(i), MergeField)
                    If (merageField IsNot Nothing) AndAlso (merageField.Prefix = "GroupStart") Then
                        foundGroupStart = True
                        Exit For
                    End If
                Next
            End While

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

        Private m_date As DateTime
        Public Property [Date]() As DateTime
            Get
                Return m_date
            End Get
            Set(ByVal value As DateTime)
                m_date = value
            End Set
        End Property
    End Class
End Namespace
