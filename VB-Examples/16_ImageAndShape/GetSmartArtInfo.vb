Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.SmartArts

Namespace GetSmartArtInfo
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a StringBuilder instance to efficiently accumulate the extracted SmartArt information
			Dim builder As New StringBuilder()

            ' Open the Word document containing SmartArt using a 'using' statement for automatic resource disposal
            Using document As Document = New Document("..\..\..\..\..\..\Data\SmartArt.docx")
                ' Iterate through every section in the document
                For Each section As Spire.Doc.Section In document.Sections
                    ' Skip the section if it or its paragraph collection is null to avoid exceptions
                    If section?.Paragraphs Is Nothing Then
                        Continue For
                    End If

                    ' Iterate through every paragraph in the current section
                    For Each paragraph As Paragraph In section.Paragraphs
                        ' Iterate through all child objects (elements) contained within the paragraph
                        For Each childObj In paragraph.ChildObjects
                            ' Check if the object is a Shape and if it specifically contains a SmartArt graphic
                            Dim shape As Spire.Doc.Fields.Shapes.Shape = TryCast(childObj, Spire.Doc.Fields.Shapes.Shape)
                            If shape IsNot Nothing AndAlso shape.HasSmartArt Then
                                ' Retrieve the SmartArt object from the shape
                                Dim smartArt As SmartArt = shape.SmartArt

                                ' Skip processing if the SmartArt object is unexpectedly null
                                If smartArt Is Nothing Then
                                    Continue For
                                End If

                                ' Append the type of the SmartArt graphic to the result string
                                builder.AppendLine($"SmartArtType£º{smartArt.SmartArtType}")

                                ' Call a helper method to extract and append background formatting details
                                ExtractSmartArtBackgroundInfo(smartArt, builder)

                                ' Call a recursive helper method to traverse nodes and extract their text and properties
                                TraverseSmartArtNodes(smartArt.Nodes, builder, 0)

                                ' Append a separator line to distinguish this SmartArt block from the next one
                                builder.AppendLine("----------------------------------------rn")
                            End If
                        Next childObj
                    Next paragraph
                Next section
            End Using

            ' Define the file path for the output text file
            Dim result As String = "GetSmartArtInfo.txt"

            ' Write the accumulated text content from the StringBuilder to the file
            File.WriteAllText(result, builder.ToString())
        End Sub

        ' Define a method to extract and append SmartArt background fill information to the StringBuilder
        Public Shared Sub ExtractSmartArtBackgroundInfo(ByVal smartArt As SmartArt, ByVal builder As StringBuilder)
            ' Return immediately if the background fill type is set to NoFill (transparent/none)
            If smartArt?.BackgroundFill.FillType = FillType.NoFill Then
                Return
            End If

            ' Convert the background fill type enum to its string representation
            Dim bgFillType As String = smartArt.BackgroundFill.FillType.ToString()

            ' Check if the background color is empty; if so, use a placeholder text, otherwise convert color to string
            Dim bgColor As String = If(smartArt.BackgroundFill.Color = Color.Empty, "No color", smartArt.BackgroundFill.Color.ToString())

            ' Append the background fill type and color information to the StringBuilder with newlines
            builder.AppendLine($"BackgroundFill_filltype£º{bgFillType}{vbLf}BackgroundFill_color£º{bgColor}")
        End Sub

        ' Define a recursive method to traverse SmartArt nodes and extract their text and properties
        Public Shared Sub TraverseSmartArtNodes(ByVal nodes As SmartArtNodeCollection, ByVal builder As StringBuilder, ByVal level As Integer)
            ' Exit the method if the node collection is null or contains no nodes
            If nodes Is Nothing OrElse nodes.Count = 0 Then Return

            ' Iterate through each node in the current collection
            For nodeIdx As Integer = 0 To nodes.Count - 1
                ' Get the current node from the collection
                Dim node As SmartArtNode = nodes(nodeIdx)

                ' Skip to the next iteration if the current node is null
                If node Is Nothing Then
                    Continue For
                End If

                ' Trim whitespace from the node text, or use a placeholder if the text is null
                Dim nodeText As String = If(node.Text IsNot Nothing, node.Text.Trim(), "Empty Text")

                ' Skip this node if the text is just a carriage return or effectively empty
                If nodeText = vbCr OrElse String.IsNullOrEmpty(nodeText) Then
                    Continue For
                End If

                ' Declare a variable to hold the prefix string based on the hierarchy level
                Dim nodePrefix As String

                ' Determine the appropriate prefix string based on the current depth level in the hierarchy
                Select Case level
                    Case 0
                        nodePrefix = "smartArt.Nodes"
                    Case 1
                        nodePrefix = "smartArt.Nodes.ChildNodes"
                    Case 2
                        nodePrefix = "smartArt.Nodes.ChildNodes.ChildNodes"
                    Case Else
                        nodePrefix = $"smartArt.Nodes.Level{level}"
                End Select

                ' Append the formatted node index and text content to the StringBuilder
                builder.AppendLine($"{nodePrefix}_{nodeIdx}£º{nodeText}")

                ' Call a helper method to extract and append specific properties of the current node
                ExtractSmartArtNodeProperties(node, builder)

                ' Recursively call this method to process any child nodes, incrementing the level counter
                TraverseSmartArtNodes(node.ChildNodes, builder, level + 1)
            Next nodeIdx
        End Sub

        ' Define a method to extract and append shape formatting properties (fill and border) of a SmartArt node
        Public Shared Sub ExtractSmartArtNodeProperties(ByVal node As SmartArtNode, ByVal builder As StringBuilder)
            ' First, call the helper method to extract font and text formatting properties
            ExtractFontProperties(node, builder)

            ' Exit the method if the node has no shape properties or the collection is empty
            If node.ShapeProperties Is Nothing OrElse node.ShapeProperties.Count = 0 Then
                Return
            End If

            ' Retrieve the first shape property object from the node to access its visual settings
            Dim shapeProps As SmartArtShapeProperties = node.ShapeProperties(0)

            ' Check if the shape has a fill type other than "NoFill" (i.e., it has a solid color or picture)
            If shapeProps?.Fill.FillType <> FillType.NoFill Then
                ' Convert the fill type enum to its string representation
                Dim nodeFillType As String = shapeProps.Fill.FillType.ToString()

                ' Determine the color string: use a specific message for pictures, otherwise get the color value
                Dim nodeColor As String = If(nodeFillType = "Picture", "(Picture type, without color acquisition)", shapeProps.Fill.Color.ToString())

                ' Append the fill type and color information to the StringBuilder with indentation
                builder.AppendLine($"{vbTab}filltype£º{nodeFillType}{vbLf}{vbTab}color£º{nodeColor}")
            End If

            ' Check if the shape's border (line format) has a fill type other than "NoFill" and a valid color
            If shapeProps.LineFormat?.Fill.FillType <> FillType.NoFill AndAlso shapeProps.LineFormat.Fill.Color <> Color.Empty Then
                ' Convert the border fill type enum to its string representation
                Dim lineFillType As String = shapeProps.LineFormat.Fill.FillType.ToString()

                ' Convert the border color to its string representation
                Dim lineColor As String = shapeProps.LineFormat.Fill.Color.ToString()

                ' Append the border fill type and color information to the StringBuilder with indentation
                builder.AppendLine($"{vbTab}line_filltype£º{lineFillType}{vbLf}{vbTab}line_color£º{lineColor}")
            End If
        End Sub

        ' Define a private helper method to extract and append font formatting properties of the node's text
        Private Shared Sub ExtractFontProperties(ByVal node As SmartArtNode, ByVal builder As StringBuilder)
            ' Return immediately if the node, its paragraphs collection, or the paragraphs themselves are null or empty
            If node?.Paragraphs Is Nothing OrElse node.Paragraphs.Count = 0 Then
                Return
            End If

            ' Get the first paragraph from the node, which typically contains the main text
            Dim paragraph = node.Paragraphs(0)

            ' Return if the paragraph's child objects collection is null or empty
            If paragraph?.ChildObjects Is Nothing OrElse paragraph.ChildObjects.Count = 0 Then
                Return
            End If

            ' Attempt to cast the first child object as a TextRange to access character formatting
            Dim textRange As Spire.Doc.Fields.TextRange = TryCast(paragraph.ChildObjects(0), Spire.Doc.Fields.TextRange)

            ' Return if the cast fails (i.e., the object is not a TextRange)
            If textRange Is Nothing Then
                Return
            End If

            ' Retrieve the CharacterFormat object which holds all font-related settings
            Dim charFormat = textRange.CharacterFormat

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
            Dim fontInfoBuilder As New StringBuilder()

            ' Check if the font name is not empty and append it if valid
            If Not String.IsNullOrEmpty(fontName) Then
                fontInfoBuilder.Append($"{vbTab}font_name£º{fontName}")
                hasValidFontInfo = True
            End If

            ' Check if the font size is greater than zero and append it if valid
            If fontSize > 0 Then
                If hasValidFontInfo Then
                    fontInfoBuilder.Append($"{vbTab}font_size£º{fontSize}pt")
                End If
                hasValidFontInfo = True
            End If

            ' Check if the font color is not empty and append it if valid
            If fontColor <> Color.Empty Then
                If hasValidFontInfo Then
                    fontInfoBuilder.Append($"{vbTab}font_color£º{fontColor}")
                End If
                hasValidFontInfo = True
            End If

            ' Append the font style information regardless of other fields
            fontInfoBuilder.Append($"{vbTab}font_style£º{fontstyle}")

            ' If any valid font information was collected, append the entire string to the main builder
            If hasValidFontInfo Then
                builder.AppendLine(fontInfoBuilder.ToString())
            End If
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
