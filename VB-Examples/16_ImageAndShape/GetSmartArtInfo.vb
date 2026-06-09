Imports System
Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.SmartArts

Namespace GetSmartArtInfo
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub
        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a StringBuilder instance to efficiently accumulate the extracted SmartArt information
            Dim builder As StringBuilder = New StringBuilder()

            ' Open the Word document containing SmartArt using a 'using' statement for automatic resource disposal
            using (Document document = New Document(@"..\..\..\..\..\..\Data\SmartArt.docx"))
                ' Iterate through every section in the document
                For Each section As Section In document.Sections
                    ' Skip the section if it or its paragraph collection is null to avoid exceptions
                    if (section?.Paragraphs  = Nothing) continue

                    ' Iterate through every paragraph in the current section
                    For Each paragraph As Paragraph In section.Paragraphs
                        ' Iterate through all child objects (elements) contained within the paragraph
                        For Each childObj As var In paragraph.ChildObjects
                            ' Check if the object is a Shape and if it specifically contains a SmartArt graphic
                            If childObj is Spire.Doc.Fields.Shapes.Shape shape && shape.HasSmartArt Then
                                ' Retrieve the SmartArt object from the shape
                                Dim smartArt As SmartArt = shape.SmartArt

                                ' Skip processing if the SmartArt object is unexpectedly null
                                if (smartArt  = Nothing) continue

                                ' Append the type of the SmartArt graphic to the result string
                                builder.AppendLine($"SmartArtType£º{smartArt.SmartArtType}")

                                ' Call a helper method to extract and append background formatting details
                                ExtractSmartArtBackgroundInfo(smartArt, builder)

                                ' Call a recursive helper method to traverse nodes and extract their text and properties
                                TraverseSmartArtNodes(smartArt.Nodes, builder, 0)

                                ' Append a separator line to distinguish this SmartArt block from the next one
                                builder.AppendLine("----------------------------------------rn")
                            End If
                        Next
                    Next
                Next
        End Sub

            ' Define the file path for the output text file
            Dim result As String = "GetSmartArtInfo.txt"

            ' Write the accumulated text content from the StringBuilder to the file
            File.WriteAllText(result, builder.ToString())
    End Class
        ' Define a method to extract and append SmartArt background fill information to the StringBuilder
        Public Sub ExtractSmartArtBackgroundInfo(ByVal smartArt As SmartArt, ByVal builder As StringBuilder)
            ' Return immediately if the background fill type is set to NoFill (transparent/none)
            If smartArt?.BackgroundFill.FillType == FillType.NoFill Then
                Return
            End If

            ' Convert the background fill type enum to its string representation
            Dim bgFillType As String = smartArt.BackgroundFill.FillType.ToString()

            ' Check if the background color is empty; if so, use a placeholder text, otherwise convert color to string
            Dim bgColor As String = smartArt.BackgroundFill.Color  = Color.Empty
                ? "No color"
                : smartArt.BackgroundFill.Color.ToString()

            ' Append the background fill type and color information to the StringBuilder with newlines
            builder.AppendLine($"BackgroundFill_filltype£º{bgFillType}\nBackgroundFill_color£º{bgColor}")
        End Sub

        ' Define a recursive method to traverse SmartArt nodes and extract their text and properties
        Public Shared Sub TraverseSmartArtNodes(ByVal nodes As SmartArtNodeCollection, ByVal builder As StringBuilder, ByVal level As Integer)
            ' Exit the method if the node collection is null or contains no nodes
            if (nodes  = Nothing  OrElse  nodes.Count  = 0) return

            ' Iterate through each node in the current collection
            Dim nodeIdx As Integer = 0
            While nodeIdx < nodes.Count
                ' Get the current node from the collection
                Dim node As SmartArtNode = nodes[nodeIdx]

                ' Skip to the next iteration if the current node is null
                if (node  = Nothing) continue

                ' Trim whitespace from the node text, or use a placeholder if the text is null
                Dim nodeText As String = node.Text  <> Nothing ? node.Text.Trim() : "Empty Text"

                ' Skip this node if the text is just a carriage return or effectively empty
                if (nodeText  = "\r"  OrElse  String.IsNullOrEmpty(nodeText)) continue

                ' Declare a variable to hold the prefix string based on the hierarchy level
                Dim nodePrefix As String

                ' Determine the appropriate prefix string based on the current depth level in the hierarchy
                Select Case level
                    Case 0
                        nodePrefix = "smartArt.Nodes"
                        Dim Select As Exit
                    Case 1
                        nodePrefix = "smartArt.Nodes.ChildNodes"
                        Dim Select As Exit
                    Case 2
                        nodePrefix = "smartArt.Nodes.ChildNodes.ChildNodes"
                        Dim Select As Exit
Case Else
                        nodePrefix = $"smartArt.Nodes.Level{level}"
                        Dim Select As Exit
                End Select

                ' Append the formatted node index and text content to the StringBuilder
                builder.AppendLine($"{nodePrefix}_{nodeIdx}£º{nodeText}")

                ' Call a helper method to extract and append specific properties of the current node
                ExtractSmartArtNodeProperties(node, builder)

                ' Recursively call this method to process any child nodes, incrementing the level counter
                TraverseSmartArtNodes(node.ChildNodes, builder, level + 1)
            Next
        End Sub
        ' Define a method to extract and append shape formatting properties (fill and border) of a SmartArt node
        Public Shared Sub ExtractSmartArtNodeProperties(ByVal node As SmartArtNode, ByVal builder As StringBuilder)
            ' First, call the helper method to extract font and text formatting properties
            ExtractFontProperties(node, builder)

            ' Exit the method if the node has no shape properties or the collection is empty
            If node.ShapeProperties == null || node.ShapeProperties.Count == 0 Then
                Return
            End If

            ' Retrieve the first shape property object from the node to access its visual settings
            Dim shapeProps As SmartArtShapeProperties = node.ShapeProperties[0]

            ' Check if the shape has a fill type other than "NoFill" (i.e., it has a solid color or picture)
            If shapeProps?.Fill.FillType != FillType.NoFill Then
                ' Convert the fill type enum to its string representation
                Dim nodeFillType As String = shapeProps.Fill.FillType.ToString()

                ' Determine the color string: use a specific message for pictures, otherwise get the color value
                Dim nodeColor As String = nodeFillType  = "Picture"
                    ? "(Picture type, without color acquisition)"
                    : shapeProps.Fill.Color.ToString()

                ' Append the fill type and color information to the StringBuilder with indentation
                builder.AppendLine($"\tfilltype£º{nodeFillType}\n\tcolor£º{nodeColor}")
            End If

            ' Check if the shape's border (line format) has a fill type other than "NoFill" and a valid color
            If shapeProps.LineFormat?.Fill.FillType != FillType.NoFill && shapeProps.LineFormat.Fill.Color != Color.Empty Then
                ' Convert the border fill type enum to its string representation
                Dim lineFillType As String = shapeProps.LineFormat.Fill.FillType.ToString()

                ' Convert the border color to its string representation
                Dim lineColor As String = shapeProps.LineFormat.Fill.Color.ToString()

                ' Append the border fill type and color information to the StringBuilder with indentation
                builder.AppendLine($"\tline_filltype£º{lineFillType}\n\tline_color£º{lineColor}")
            End If
        End Sub

        ' Define a private helper method to extract and append font formatting properties of the node's text
        Private Shared Sub ExtractFontProperties(ByVal node As SmartArtNode, ByVal builder As StringBuilder)
            ' Return immediately if the node, its paragraphs collection, or the paragraphs themselves are null or empty
            If node?.Paragraphs == null || node.Paragraphs.Count == 0 Then
                Return

            ' Get the first paragraph from the node, which typically contains the main text
            Dim paragraph As var = node.Paragraphs[0]

            ' Return if the paragraph's child objects collection is null or empty
            If paragraph?.ChildObjects == null || paragraph.ChildObjects.Count == 0 Then
                Return

            ' Attempt to cast the first child object as a TextRange to access character formatting
            Dim textRange As var = paragraph.ChildObjects[0] as Spire.Doc.Fields.TextRange

            ' Return if the cast fails (i.e., the object is not a TextRange)
            If textRange == null Then
                Return

            ' Retrieve the CharacterFormat object which holds all font-related settings
            Dim charFormat As var = textRange.CharacterFormat

            ' Extract the font name from the character format
            Dim fontName As String = charFormat.FontName

            ' Extract the font size from the character format
            Dim fontSize As Single = charFormat.FontSize

            ' Extract the text color from the character format
            Dim fontColor As Color = charFormat.TextColor

            ' Extract the font style (bold, italic, etc.) from the character format
            Dim fontstyle As Spire.Doc.Publics.Drawing.FontStyle = charFormat.FontStyle

            ' Initialize a flag to track whether any valid font information was found
            Dim hasValidFontInfo As Boolean = False

            ' Create a new StringBuilder to accumulate the font property strings efficiently
            Dim fontInfoBuilder As StringBuilder = New StringBuilder()

            ' Check if the font name is not empty and append it if valid
            If !string.IsNullOrEmpty(fontName) Then
                fontInfoBuilder.Append($"\tfont_name£º{fontName}")
                hasValidFontInfo = True
            End If

            ' Check if the font size is greater than zero and append it if valid
            If fontSize > 0 Then
                If hasValidFontInfo Then
                    fontInfoBuilder.Append($"\tfont_size£º{fontSize}pt")
                hasValidFontInfo = True
                End If

            ' Check if the font color is not empty and append it if valid
            If fontColor != Color.Empty Then
                If hasValidFontInfo Then
                    fontInfoBuilder.Append($"\tfont_color£º{fontColor}")
                hasValidFontInfo = True
                End If

            ' Append the font style information regardless of other fields
            fontInfoBuilder.Append($"\tfont_style£º{fontstyle}")

            ' If any valid font information was collected, append the entire string to the main builder
            If hasValidFontInfo Then
                builder.AppendLine(fontInfoBuilder.ToString())
            End If
            End If
        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

            End If
            End If
