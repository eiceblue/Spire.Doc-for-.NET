Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO

Namespace CheckDigitalSignature
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
				Dim hasDigitalSignature As Boolean = Document.HasDigitalSignature("..\..\..\..\..\..\Data\CheckDigitalSignature.docx")

				' Use a switch statement to determine the file format and update the fileFormat string accordingly
				If hasDigitalSignature Then
					MessageBox.Show("This Word document has digital signature")
				Else
					MessageBox.Show("This Word document has not digital signature")
				End If
		End Sub
	End Class
End Namespace
