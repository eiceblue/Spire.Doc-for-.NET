Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO
Imports Spire.Doc.Fields
Imports System.Data.OleDb
Namespace DocAndDataBase
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Set the input database file path
			Dim inputDataBase As String = "..\..\..\..\..\..\Data\demo.mdb"

			' Set the input folder path
			Dim inputFolder As String = "..\..\..\..\..\..\Data\"

			' Specify the file name to be used as a template
			Dim fileName As String = "Template.docx"

			' Define the connection string for the OleDbConnection
			Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & inputDataBase

			' Create a new OleDbConnection using the connection string and open the connection
			Dim connection As New OleDbConnection(connString)
			connection.Open()

			' Store the specified document in the database
			StoreToDatabase(inputFolder & fileName, connection)

			' Read the document from the database into a Document object
			Dim dbDoc As Document = ReadFromDatabase(fileName, connection)

			' Specify the new file name for the retrieved document
			Dim newFileName As String = "DocAndDataBase_out.docx"

			' Save the retrieved document to a new file in Docx format
			dbDoc.SaveToFile(newFileName, FileFormat.Docx)

			' Delete the document from the database
			DeleteFromDatabase(fileName, connection)

			' Close the connection to the database
			connection.Close()
			
			' Dispose of the Document object to release resources
			dbDoc.Dispose()

			'Launching the MS Word file.
			WordDocViewer("DocAndDataBase_out.docx")
		End Sub
		
		'Store document to database 
		Public Shared Sub StoreToDatabase(ByVal input As String, ByVal connection As OleDbConnection)
			' Create a new Document object using the specified input file
			Dim doc As New Document(input)

			' Create a new memory stream to store the document content
			Dim stream As New MemoryStream()

			' Save the document to the memory stream in Docx format
			doc.SaveToStream(stream, FileFormat.Docx)

			' Get the file name from the input path
			Dim fileName As String = Path.GetFileName(input)

			' Define the SQL command to insert the document content into the database
			Dim commandString As String = "INSERT INTO Documents (FileName, FileContent) VALUES('" & fileName & "', @Doc)"

			' Create a new OleDbCommand with the command string and connection
			Dim command As New OleDbCommand(commandString, connection)

			' Add the document content as a parameter to the command
			command.Parameters.AddWithValue("Doc", stream.ToArray())

			' Execute the SQL command to insert the document into the database
			command.ExecuteNonQuery()
		End Sub

		' Read document from database 
		Public Shared Function ReadFromDatabase(ByVal fileName As String, ByVal mConnection As OleDbConnection) As Document
			' Define the SQL command to retrieve the document content from the database
		Dim commandString As String = "SELECT * FROM Documents WHERE FileName='" & fileName & "'"

		' Create a new OleDbCommand with the command string and connection
		Dim command As New OleDbCommand(commandString, mConnection)

		' Create a new OleDbDataAdapter to fill a DataTable with the result of the command
		Dim adapter As New OleDbDataAdapter(command)

		' Create a new DataTable to store the retrieved data
		Dim dataTable As New DataTable()

		' Fill the DataTable with the result of the SQL query
		adapter.Fill(dataTable)

		' Check if the DataTable contains any rows
		If dataTable.Rows.Count = 0 Then
			' Throw an exception if no matching record is found in the database
			Throw New ArgumentException(String.Format("Could not find any record matching the document ""{0}"" in the database.", fileName))
		End If

		' Retrieve the byte array representing the document content from the DataTable
		Dim buffer() As Byte = CType(dataTable.Rows(0)("FileContent"), Byte())

		' Create a new memory stream using the retrieved byte array
		Dim newStream As New MemoryStream(buffer)

		' Create a new Document object using the content from the memory stream
		Dim doc As New Document(newStream)

		' Return the retrieved Document object
		Return doc
		End Function

		' Delete document from database 
		Public Shared Sub DeleteFromDatabase(ByVal fileName As String, ByVal mConnection As OleDbConnection)
			 ' Define the SQL command to delete the specified document from the database
			Dim commandString As String = "DELETE * FROM Documents WHERE FileName='" & fileName & "'"

			' Create a new OleDbCommand with the command string and connection
			Dim command As New OleDbCommand(commandString, mConnection)

			' Execute the SQL command to delete the document from the database
			command.ExecuteNonQuery()
		End Sub
		
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
