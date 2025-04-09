Imports System.Collections
Imports Spire.Doc

Namespace NestedMailMerage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a list to store DictionaryEntry objects
			Dim list As New List(Of DictionaryEntry)()

			'Create a DataSet object
			Dim dsData As New DataSet()

			'Read XML data into the DataSet
			dsData.ReadXml("..\..\..\..\..\..\Data\Orders.xml")

			'Create a Document object
			Dim document As New Document()

			'Load a Word document from file
			document.LoadFromFile("..\..\..\..\..\..\Data\NestedMailMerge.doc")

			'Create a DictionaryEntry for "Customer" with an empty value and add it to the list
			Dim dictionaryEntry As New DictionaryEntry("Customer", String.Empty)
			list.Add(dictionaryEntry)

			'Create a DictionaryEntry for "Order" with a nested region condition and add it to the list
			dictionaryEntry = New DictionaryEntry("Order", "Customer_Id = %Customer.Customer_Id%")
			list.Add(dictionaryEntry)

			'Execute mail merge with nested regions using the DataSet and list of DictionaryEntry objects
			document.MailMerge.ExecuteWidthNestedRegion(dsData, list)

			'Save the merged document to a file 
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			'Dispose the document
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("Sample.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
