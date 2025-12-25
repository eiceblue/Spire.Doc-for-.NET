Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Fields.Shape
Imports Spire.Doc.Fields.Shapes
Namespace AdjustRoundRectangleCornerRadius
    Partial Public Class Form1
        Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            ' Load the existing Word document
            Dim document As Document = New Document("..\..\..\..\..\..\Data\AdjustRoundRectangleCornerRadius.docx")

            ' Get the first section of the document
            Dim section As Section = document.Sections(0)

            ' Iterate through all child objects in the section's body
            For Each obj As DocumentObject In section.Body.ChildObjects
                ' Check if the current object is a paragraph
                If TypeOf obj Is Paragraph Then
                    ' Cast the object to a Paragraph
                    Dim paragraph As Paragraph = DirectCast(obj, Paragraph)

                    ' Iterate through all child objects within the paragraph
                    For Each docObj As DocumentObject In paragraph.ChildObjects
                        ' Check if the current child object is a Shape
                        If TypeOf docObj Is Shape Then
                            ' Cast the child object to a ShapeObject
                            Dim shape As ShapeObject = DirectCast(docObj, ShapeObject)

                            ' Check if the shape type is a Round Rectangle
                            If shape.ShapeType = ShapeType.RoundRectangle Then
                                ' Get the current corner radius of the round rectangle
                                Dim cornerRadius As Double = shape.AdjustHandles.GetRoundRectangleCornerRadius()

                                ' Adjust the corner radius of the round rectangle to 20
                                shape.AdjustHandles.AdjustRoundRectangle(20)
                            End If
                        End If
                    Next
                End If
            Next

            ' Define the output file name
            Dim result As String = "AdjustRoundRectangleCornerRadius-result.docx"

            ' Save the modified document to a new file in Docx 2016 format
            document.SaveToFile(result, FileFormat.Docx2016)

            ' Dispose of the document object to free resources
            document.Dispose()

            'Launching the MS Word file.
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
