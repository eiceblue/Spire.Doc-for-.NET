Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.SmartArts

Namespace ModifySmartArt
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new instance of the Document class to represent a Word document
            Dim document As Document = New Document()

            ' Load an existing Word document containing SmartArt from the specified file path
            document.LoadFromFile(@"..\..\..\..\..\..\Data\SmartArt.docx")

            ' Get the first section of the loaded document
            Dim section As Section = document.Sections[0]

            ' Get the first paragraph from the first section, which is expected to contain the SmartArt
            Dim paragraph As Paragraph = section.Paragraphs[0]

            ' Retrieve the first child object from the paragraph and cast it as a Shape (the SmartArt container)
            Dim shape1 As Spire.Doc.Fields.Shapes.Shape = paragraph.ChildObjects[0] as Spire.Doc.Fields.Shapes.Shape

            ' Access the SmartArt object from the retrieved shape
            Dim smartArt As SmartArt = shape1.SmartArt

            ' Set the background fill type of the SmartArt to a solid color
            smartArt.BackgroundFill.FillType = FillType.Solid

            ' Set the background color of the SmartArt to a light orange/peach color using ARGB values
            smartArt.BackgroundFill.Color = Color.FromArgb(255, 242, 169, 132)

            ' Get the first node (main shape) of the SmartArt graphic
            Dim node As SmartArtNode = smartArt.Nodes[0]

            ' Set the text content of the first node to "Goals"
            node.Text = "Goals"

            ' Get the shape properties of the first node to customize its appearance
            Dim shape As SmartArtShapeProperties = node.ShapeProperties[0]

            ' Set the fill type of the shape to a solid color
            shape.Fill.FillType = FillType.Solid

            ' Set the fill color of the shape to a purple color using ARGB values
            shape.Fill.Color = Color.FromArgb(255, 160, 43, 147)

            ' Set the fill type of the shape's border (line format) to a solid color
            shape.LineFormat.Fill.FillType = FillType.Solid

            ' Set the border color of the shape to the same purple color
            shape.LineFormat.Fill.Color = Color.FromArgb(255, 160, 43, 147)

            ' Get the first child node of the "Goals" node
            Dim childNode As SmartArtNode = node.ChildNodes[0]

            ' Set the text content of the child node to a descriptive sentence
            childNode.Text = "Set clear goals to the team."

            ' Set the border fill type of the child node's shape to a solid color
            childNode.ShapeProperties[0].LineFormat.Fill.FillType = FillType.Solid

            ' Set the border color of the child node's shape to the same purple color
            childNode.ShapeProperties[0].LineFormat.Fill.Color = Color.FromArgb(255, 160, 43, 147)

            ' Get the second main node of the SmartArt graphic
            node = smartArt.Nodes[1]

            ' Set the text content of the second node to "Progress"
            node.Text = "Progress"

            ' Get the third main node of the SmartArt graphic
            node = smartArt.Nodes[2]

            ' Set the text content of the third node to "Result"
            node.Text = "Result"

            ' Get the shape properties of the third node to customize its appearance
            shape = node.ShapeProperties[0]

            ' Set the fill type of the third node's shape to a solid color
            shape.Fill.FillType = FillType.Solid

            ' Set the fill color of the third node's shape to a green color using ARGB values
            shape.Fill.Color = Color.FromArgb(255, 78, 167, 46)

            ' Set the fill type again (redundant but present in original code)
            shape.Fill.FillType = FillType.Solid

            ' Set the border color of the third node's shape to the same green color
            shape.LineFormat.Fill.Color = Color.FromArgb(255, 78, 167, 46)

            ' Set the border fill type of the first child node under the "Result" node to a solid color
            node.ChildNodes[0].ShapeProperties[0].LineFormat.Fill.FillType = FillType.Solid

            ' Set the border color of the first child node under the "Result" node to the same green color
            node.ChildNodes[0].ShapeProperties[0].LineFormat.Fill.Color = Color.FromArgb(255, 78, 167, 46)

            ' Define the output file name for saving the modified document
            Dim result As String = "ModifySmartArt.docx"

            ' Save the modified document to a file in Docx2016 format
            document.SaveToFile(result, FileFormat.Docx2016)

            ' Close the document to release any file handles or resources
            document.Close()

            ' Dispose of the document object to free up system memory
            document.Dispose()

            WordDocViewer(result)
        End Sub
      
        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

    End Class
End Namespace
