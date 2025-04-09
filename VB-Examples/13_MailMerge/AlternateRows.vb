Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO
Imports Spire.Doc.Fields
Imports System.Data.OleDb
Imports System.Linq
Imports Spire.Doc.Reporting
Imports System.Collections
Namespace AlternateRows
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

			'Create a Word document
			Dim doc As New Document()

			'Load a mail merge template file
			doc.LoadFromFile(input)

			AddHandler doc.MailMerge.MergeField, AddressOf MailMerge_MergeField

			'Fill mergedField with data from dataTable
			doc.MailMerge.ExecuteWidthRegion(orderTable)

			'Save to file
			Dim result As String = "AlternateRows_out.doc"
			doc.SaveToFile(result, FileFormat.Doc)

			' Dispose the document
			doc.Dispose()
			
			WordViewer(result)
		End Sub
		Private rowIndex As Integer = 0
		Private Sub MailMerge_MergeField(ByVal sender As Object, ByVal args As MergeFieldEventArgs)
			' Catch the beginning of a new row.
			If args.CurrentMergeField.FieldName.Equals("Name") Then
				' Set the color depending on whether the row number is even or odd.
				Dim rowColor As Color
				If rowIndex Mod 2 = 0 Then
					rowColor = Color.FromArgb(215, 227, 235)
				Else
					rowColor = Color.FromArgb(240, 242, 242)
				End If

				' Get the owner cell
				Dim cell As TableCell = TryCast(args.CurrentMergeField.OwnerParagraph.Owner, TableCell)

				' Get the owner row
				Dim row As TableRow = cell.OwnerRow

				'Set the back color
				For i As Integer = 0 To row.Cells.Count - 1
					row.Cells(i).CellFormat.Shading.BackgroundPatternColor =rowColor
				Next i

				rowIndex += 1
			End If
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
			adapter.Fill(dataTable)
			dataTable.TableName = "Country"

			' Close the connection
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
