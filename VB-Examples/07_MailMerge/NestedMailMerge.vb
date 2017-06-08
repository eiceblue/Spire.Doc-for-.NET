Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.Collections

Namespace NestedMailMerage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim list_Renamed As New List(Of DictionaryEntry)()
			Dim dsData As New DataSet()

			dsData.ReadXml("..\..\..\..\..\..\Data\Orders.xml")

			'Create word document
			Dim document_Renamed As New Document()
			document_Renamed.LoadFromFile("..\..\..\..\..\..\Data\Invoice.doc")

			Dim dictionaryEntry_Renamed As New DictionaryEntry("Customer", String.Empty)
			list_Renamed.Add(dictionaryEntry_Renamed)

			dictionaryEntry_Renamed = New DictionaryEntry("Order", "Customer_Id = %Customer.Customer_Id%")
			list_Renamed.Add(dictionaryEntry_Renamed)

			document_Renamed.MailMerge.ExecuteWidthNestedRegion(dsData, list_Renamed)

			'Save doc file.
			document_Renamed.SaveToFile("Sample.doc", FileFormat.Doc)

			'Launching the MS Word file.
			WordDocViewer("Sample.doc")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
