Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO

Namespace GetTablePosition
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new Document object
			Dim document As New Document()

			' Load an existing Word document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\TableSample-Az.docx")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Get the first table in the section
			Dim table As Table = TryCast(section.Tables(0), Table)

			' Create a StringBuilder to store the output content
			Dim stringBuilder As New StringBuilder()

			' Check if text wrapping is enabled around the table
			If table.Format.WrapTextAround Then
				' Get the positioning information for the table
				Dim position As TablePositioning = table.Format.Positioning

				' Append horizontal positioning information to the output content
				stringBuilder.AppendLine("Horizontal:")
				stringBuilder.AppendLine("Position: " & position.HorizPosition & " pt")
				stringBuilder.AppendLine("Absolute Position: " & position.HorizPositionAbs & ", Relative to: " & position.HorizRelationTo)
				stringBuilder.AppendLine()

				' Append vertical positioning information to the output content
				stringBuilder.AppendLine("Vertical:")
				stringBuilder.AppendLine("Position: " & position.VertPosition & " pt")
				stringBuilder.AppendLine("Absolute Position: " & position.VertPositionAbs & ", Relative to: " & position.VertRelationTo)
				stringBuilder.AppendLine()

				' Append distance from surrounding text information to the output content
				stringBuilder.AppendLine("Distance from surrounding text:")
				stringBuilder.AppendLine("Top: " & position.DistanceFromTop & " pt, Left: " & position.DistanceFromLeft & " pt")
				stringBuilder.AppendLine("Bottom: " & position.DistanceFromBottom & " pt, Right: " & position.DistanceFromRight & " pt")
			End If

			' Specify the output file path
			Dim result As String = "GetTablePosition_out.txt"

			' Write the output content to the output file
			File.WriteAllText(result, stringBuilder.ToString())

			' Dispose of the document object to free up resources
			document.Dispose()

			'Launching the Word file.
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
