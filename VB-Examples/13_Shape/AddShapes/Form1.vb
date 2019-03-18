Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddShapes
    Partial Public Class Form1
        Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            Dim doc As New Document()
            Dim sec As Section = doc.AddSection()
            Dim para As Paragraph = sec.AddParagraph()
            Dim x As Integer = 60, y As Integer = 40, lineCount As Integer = 0
            For i As Integer = 1 To 20
                If lineCount > 0 AndAlso lineCount Mod 8 = 0 Then
                    para.AppendBreak(BreakType.PageBreak)
                    x = 60
                    y = 40
                    lineCount = 0
                End If
                'add shape and set its size and position
                Dim shape As ShapeObject = para.AppendShape(50, 50, CType(i, ShapeType))
                shape.HorizontalOrigin = HorizontalOrigin.Page
                shape.HorizontalPosition = x
                shape.VerticalOrigin = VerticalOrigin.Page
                shape.VerticalPosition = y + 50
                x = x + CInt(shape.Width) + 50
                If i > 0 AndAlso i Mod 5 = 0 Then
                    y = y + CInt(shape.Height) + 120
                    lineCount += 1
                    x = 60
                End If

            Next i
            'Save docx file
            doc.SaveToFile("AddShape.docx", FileFormat.Docx)

            'Launch Word file.
            WordDocViewer("AddShape.docx")
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                Process.Start(fileName)
            Catch
            End Try
        End Sub
    End Class
End Namespace