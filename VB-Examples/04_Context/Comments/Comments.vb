Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace Comments
    Partial Public Class Form1
        Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            'Open a blank word document as template
            Dim document_Renamed As New Document("..\..\..\..\..\..\Data\Blank.doc")

            InsertComments(document_Renamed.Sections(0))

            'Save doc file.
            document_Renamed.SaveToFile("Sample.doc", FileFormat.Doc)

            'Launching the MS Word file.
            WordDocViewer("Sample.doc")


        End Sub

        Private Sub InsertComments(ByVal section As Section)
            'title
            Dim paragraph As Paragraph = Nothing
            If section.Paragraphs.Count > 0 Then
                paragraph = section.Paragraphs(0)
            Else
                paragraph = section.AddParagraph()
            End If

            Dim title As TextRange = paragraph.AppendText("Summary of Science")
            title.CharacterFormat.Bold = True
            title.CharacterFormat.FontName = "Arial"
            title.CharacterFormat.FontSize = 14
            paragraph.Format.HorizontalAlignment _
                = Spire.Doc.Documents.HorizontalAlignment.Center
            paragraph.Format.AfterSpacing = 10

            'style
            Dim style1 As ParagraphStyle = New ParagraphStyle(section.Document)
            style1.Name = "style1"
            style1.CharacterFormat.FontName = "Arial"
            style1.CharacterFormat.FontSize = 9
            style1.ParagraphFormat.LineSpacing = 1.5F * 12.0F
            style1.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple
            section.Document.Styles.Add(style1)

            Dim style2 As ParagraphStyle = New ParagraphStyle(section.Document)
            style2.Name = "style2"
            style2.ApplyBaseStyle(style1.Name)
            style2.CharacterFormat.Font = New Font("Arial", 10.0F)
            section.Document.Styles.Add(style2)

            paragraph = section.AddParagraph()
            paragraph.AppendText("(All text and pictures are from ")
            Dim text As TextRange = paragraph.AppendText("Wikipedia")

            'Comment Wikipedia, adding url for it.
            Dim comment1 As Comment = paragraph.AppendComment("http://en.wikipedia.org/wiki/Science")
            comment1.AddItem(text)
            comment1.Format.Author = "Harry Hu"
            comment1.Format.Initial = "HH"
            paragraph.AppendText(", the free encyclopedia)")
            paragraph.ApplyStyle(style1.Name)

            Dim paragraph1 As Paragraph = section.AddParagraph()
            Dim str1 As String _
                = "Science (from the Latin scientia, meaning ""knowledge"") " _
                + "is an enterprise that builds and organizes knowledge in the form " _
                + "of testable explanations and predictions about the natural world. " _
                + "An older meaning still in use today is that of Aristotle, " _
                + "for whom scientific knowledge was a body of reliable knowledge " _
                + "that can be logically and rationally explained " _
                + "(see ""History and etymology"" section below)."
            paragraph1.AppendText(str1)

            'Simple comment
            Dim comment2 As Comment = paragraph1.AppendComment("Not given in this file.")
            comment2.Format.Author = "Harry Hu"
            comment2.Format.Initial = "HH"

            'Insert a picture in the right of the paragraph1
            Dim picture As DocPicture _
                = paragraph1.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Wikipedia_Science.png"))
            picture.TextWrappingStyle = TextWrappingStyle.Square
            picture.TextWrappingType = TextWrappingType.Left
            picture.VerticalOrigin = VerticalOrigin.Paragraph
            picture.VerticalPosition = 0
            picture.HorizontalOrigin = HorizontalOrigin.Column
            picture.HorizontalAlignment = ShapeHorizontalAlignment.Right

            paragraph1.ApplyStyle(style2.Name)

            Dim paragraph2 As Paragraph = section.AddParagraph()
            Dim str2 As String _
                = "Since classical antiquity science as a type of knowledge was closely linked " _
                + "to philosophy, the way of life dedicated to discovering such knowledge. " _
                + "And into early modern times the two words, ""science"" and ""philosophy"", " _
                + "were sometimes used interchangeably in the English language. " _
                + "By the 17th century, ""natural philosophy"" " _
                + "(which is today called ""natural science"") could be considered separately " _
                + "from ""philosophy"" in general. But ""science"" continued to also be used " _
                + "in a broad sense denoting reliable knowledge about a topic, in the same way " _
                + "it is still used in modern terms such as library science or political science."
            paragraph2.AppendText(str2)
            paragraph2.ApplyStyle(style2.Name)

            Dim paragraph3 As Paragraph = section.AddParagraph()
            Dim str3 As String _
                = "The more narrow sense of ""science"" that is common today developed as a part " _
                + "of science became a distinct enterprise of defining ""laws of nature"", " _
                + "based on early examples such as Kepler's laws, Galileo's laws, and Newton's " _
                + "laws of motion. In this period it became more common to refer to natural " _
                + "philosophy as  ""natural science"". Over the course of the 19th century, the word " _
                + """science"" became increasingly associated with the disciplined study of the " _
                + "natural world including physics, chemistry, geology and biology. This sometimes " _
                + "left the study of human thought and society in a linguistic limbo, which was " _
                + "resolved by classifying these areas of academic study as social science. " _
                + "Similarly, several other major areas of disciplined study and knowledge " _
                + "exist today under the general rubric of ""science"", such as formal science " _
                + "and applied science."
            paragraph3.AppendText(str3)
            paragraph3.ApplyStyle(style2.Name)
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                Process.Start(fileName)
            Catch
            End Try
        End Sub

    End Class
End Namespace
