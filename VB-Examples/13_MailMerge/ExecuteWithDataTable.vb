Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO
Imports Spire.Doc.Fields
Imports System.Data.OleDb
Namespace ExecuteWithDataTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim inputDataBase As String = "..\..\..\..\..\..\Data\demo.mdb"
			Dim input As String = "..\..\..\..\..\..\Data\ExecuteWithDataTable.doc"

			' Get a dataTable
			Dim orderTable As DataTable = GetCountryDataTable(inputDataBase)

			'Create a Document 
			Dim doc As New Document()

			'Load a mail merge template file
			doc.LoadFromFile(input)

			'Fill mergedField with data from dataTable
			doc.MailMerge.ExecuteWidthRegion(orderTable)

			'Save to file
			Dim result As String = "ExecuteWithDataTable_out.doc"
			doc.SaveToFile(result, FileFormat.Doc)

			'Dispose the document
			doc.Dispose()
			WordViewer(result)
		End Sub
		Private Function GetCountryDataTable(ByVal inputDataBase As String) As DataTable
			' Open a database connection
			Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & inputDataBase
			Dim connection As New OleDbConnection(connString)
			connection.Open()

			' Create the SQL command.
			Dim commandString As String = "SELECT * FROM Country"
			Dim command As New OleDbCommand(commandString, connection)

			' Create the data adapter.
			Dim adapter As New OleDbDataAdapter(command)

			' Fill the results from the database into a DataTable.
			Dim dataTable As New DataTable()

			' Fill the data table
			adapter.Fill(dataTable)
			dataTable.TableName = "Country"

			'Close the connection
			connection.Close()

			Return dataTable
		End Function
		Private Sub WordViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
