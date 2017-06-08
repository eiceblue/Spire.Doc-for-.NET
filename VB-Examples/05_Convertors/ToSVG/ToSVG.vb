Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ToSVG
    Partial Public Class Form1
        Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            'Create word document
            Dim document As New Document()
            document.LoadFromFile("..\..\..\..\..\..\Data\Word.doc")
            document.SaveToFile("Sample.svg", FileFormat.SVG)
        End Sub

    End Class
End Namespace
