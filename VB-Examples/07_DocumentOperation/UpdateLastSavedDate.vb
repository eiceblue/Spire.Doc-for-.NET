Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace UpdateLastSavedDate
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Define the input file path
			Dim inputFile As String = "../../../../../../Data/Template.docx"

			' Define the output file path
			Dim resultFile As String = "UpdateLastSavedDate_out.docx"

			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the specified document file
			document.LoadFromFile(inputFile)

			' Set the LastSaveDate property of the builtin document properties to the current time converted to Greenwich Mean Time
			document.BuiltinDocumentProperties.LastSaveDate = LocalTimeToGreenwishTime(Date.Now)

			' Save the modified document to a new file
			document.SaveToFile(resultFile, FileFormat.Docx)

			' Dispose of the document object
			document.Dispose()
		
			WordDocViewer(resultFile)

		End Sub

		
		' Function to convert local time to Greenwich Mean Time
		Public Shared Function LocalTimeToGreenwishTime(ByVal localTime As Date) As Date
			' Get the local time zone
			Dim localTimeZone As TimeZone = TimeZone.CurrentTimeZone
			' Get the time difference between local time and UTC
			Dim timeSpan As TimeSpan = localTimeZone.GetUtcOffset(localTime)
			' Subtract the time difference from the local time to get Greenwich Mean Time
			Dim greenwishTime As Date = localTime.Subtract(timeSpan)
			Return greenwishTime
		End Function

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
