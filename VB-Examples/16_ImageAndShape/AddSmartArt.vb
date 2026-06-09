Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.SmartArts

Namespace AddSmartArt
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        private const Integer SmartArtDefaultWidth = 432
        private const Integer SmartArtDefaultHeight = 252
        private const Single TitleFontSize = 28f
        private const Single DefaultNodeFontSize = 15f
        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new instance of the Document class to represent a Word document
            Dim document As Document = New Document()

            ' Add a new section to the document, which serves as a container for content like paragraphs and SmartArt
            Dim section As Section = document.AddSection()

            ' Initialize a list to store various types of SmartArt graphics that will be added to the document
            Dim smartArtTypes As List<SmartArtType> = New List<SmartArtType>()

            ' Add a vertical chevron list SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.VerticalChevronList)

            ' Add a square accent list SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.SquareAccentList)

            ' Add an alternating hexagons SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.AlternatingHexagons)

            ' Add a horizontal bullet list SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.HorizontalBulletList)

            ' Add a segmented process SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.SegmentedProcess)

            ' Add a vertical bending process SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.VerticalBendingProcess)

            ' Add a step down process SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.StepDownProcess)

            ' Add a circle accent timeline SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.CircleAccentTimeLine)

            ' Add a block cycle SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.BlockCycle)

            ' Add a segmented cycle SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.SegmentedCycle)

            ' Iterate through each SmartArt type in the list to create and configure SmartArt graphics
            For Each smartArtType As SmartArtType In smartArtTypes
                ' Call a helper method to create a title paragraph with the name of the current SmartArt type
                CreateTitleParagraph(section, smartArtType.ToString())

                ' Add a new paragraph to the section which will contain the SmartArt graphic
                Dim paragraph As var = section.AddParagraph()

                ' Call a helper method to create and insert the SmartArt graphic into the paragraph
                Dim smartArt As var = CreateSmartArt(paragraph, smartArtType)

                ' Get the shape properties of the first node's first shape to customize its appearance
                Dim shapeSmartArt As SmartArtShapeProperties = smartArt.Nodes[0].ShapeProperties[0]

                ' Set the fill type of the shape to a solid color
                shapeSmartArt.Fill.FillType = FillType.Solid

                ' Set the fill color of the shape to orange using ARGB values (255, 165, 0)
                shapeSmartArt.Fill.Color = Color.FromArgb(255, 255, 165, 0)

                ' Set the text and font size for the first node of the SmartArt graphic
                SetSmartArtNodeText(smartArt.Nodes[0], "TextTest_1", 15f)

                ' Add a child node to the first node with specified font size and text
                AddSmartArtChildNode(smartArt.Nodes[0], 15f, "ChildNodeTest_1.")

                ' Set the text and a larger font size for the second node of the SmartArt graphic
                SetSmartArtNodeText(smartArt.Nodes[1], "TextTest_2", 25f)

                ' Add a child node to the second node with specified font size and text
                AddSmartArtChildNode(smartArt.Nodes[1], 15f, "ChildNodeTest_2.")
            Next

            ' Define the file path and name for saving the generated Word document
            Dim result As String = "AddSmartArt.docx"

            ' Save the document to a file in Docx2016 format
            document.SaveToFile(result, FileFormat.Docx2016)

            ' Close the document to release any file handles or resources
            document.Close()

            ' Dispose of the document object to free up system memory
            document.Dispose()

            WordDocViewer(result)
        End Sub
        ' Define a method to create and return a title paragraph with specified text and formatting
        private Spire.Doc.Documents.Paragraph CreateTitleParagraph(Section section, String titleText, Single fontSize = TitleFontSize)
            ' Add a new paragraph to the given section
            Dim paragraph As var = section.AddParagraph()

            ' Set the horizontal alignment of the paragraph to center
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

            ' Append the title text to the paragraph and get the TextRange object for formatting
            Dim textRange As var = paragraph.AppendText(titleText)

            ' Set the font size of the title text
            textRange.CharacterFormat.FontSize = fontSize

            ' Set the font name of the title text to Times New Roman
            textRange.CharacterFormat.FontName = "Times New Roman"

            ' Add two empty paragraphs after the title to create vertical spacing
            section.AddParagraph()
            section.AddParagraph()

            ' Return the created title paragraph
            Return paragraph
    End Class

        ' Define a method to create and insert a SmartArt graphic into a paragraph
        Private Function CreateSmartArt(ByVal paragraph As Spire.Doc.Documents.Paragraph, ByVal smartArtType As SmartArtType, ByVal SmartArtDefaultWidth As int width =, ByVal SmartArtDefaultHeight As int height =) As SmartArt
            ' Set the horizontal alignment of the paragraph containing the SmartArt to center
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

            ' Append a SmartArt shape of the specified type and dimensions to the paragraph
            Dim shape As Spire.Doc.Fields.Shapes.Shape = paragraph.AppendSmartArt(smartArtType, width, height)

            ' Return the SmartArt object from the created shape
            Return shape.SmartArt
        End Function

        ' Define a method to set the text and font size for a SmartArt node
        Private Sub SetSmartArtNodeText(ByVal node As SmartArtNode, ByVal text As String, ByVal DefaultNodeFontSize As float fontSize =)
            ' Exit the method early if the node is null or the text is empty
            if (node  = Nothing  OrElse  String.IsNullOrEmpty(text)) return

            ' Set the text content of the SmartArt node
            node.Text = text

            ' Check if the node has paragraphs and the first paragraph has child objects (like TextRange)
            If node.Paragraphs.Count > 0 && node.Paragraphs[0].ChildObjects.Count > 0 Then
                ' Cast the first child object to TextRange and set its font size
                ((Spire.Doc.Fields.TextRange)node.Paragraphs[0].ChildObjects[0]).CharacterFormat.FontSize = fontSize
            End If
        End Sub

        ' Define a method to add child nodes to a parent SmartArt node and set their text
        Private Sub AddSmartArtChildNode(ByVal parentNode As SmartArtNode, ByVal fontSize As Single, ParamArray childTexts As string[])
            ' Exit the method early if the parent node is null, or the child texts array is null or empty
            if (parentNode  = Nothing  OrElse  childTexts  = Nothing  OrElse  childTexts.Length  = 0) return

            ' Ensure the parent node has enough child nodes to accommodate all the provided text strings
            While parentNode.ChildNodes.Count < childTexts.Length
                ' Add a new child node until the count matches the number of text strings
                parentNode.ChildNodes.Add()
            End While

            ' Iterate through the child text strings and corresponding child nodes
            Dim i As Integer = 0
            While i < childTexts.Length && i < parentNode.ChildNodes.Count
                ' Call SetSmartArtNodeText to set the text and font size for each child node
                SetSmartArtNodeText(parentNode.ChildNodes[i], childTexts[i], fontSize)
            Next
        End Sub
        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

End Namespace
