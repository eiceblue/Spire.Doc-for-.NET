Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace CreateTableOfContentByDefault
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new document
			Dim doc As New Document()

			' Add a section to the document
			Dim section As Section = doc.AddSection()

			' Add a paragraph to the section and append a table of contents (TOC)
			Dim para As Paragraph = section.AddParagraph()
			para.AppendTOC(1, 3)

			' Add a paragraph to the section 
			Dim par As Paragraph = section.AddParagraph()
			Dim tr As TextRange = par.AppendText("Flowers")
			tr.CharacterFormat.FontSize = 30
			par.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			' Add a paragraph to the section 
			Dim para1 As Paragraph = section.AddParagraph()
			para1.AppendText("Ornithogalum")

			' Apply the "Heading1" style
			para1.ApplyStyle(BuiltinStyle.Heading1)

			' Add a paragraph to the section and insert an image 
			para1 = section.AddParagraph()
			Dim picture As DocPicture = para1.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Ornithogalum.jpg"))
			picture.TextWrappingStyle = TextWrappingStyle.Square
			para1.AppendText("Ornithogalum is a genus of perennial plants mostly native to southern Europe and southern Africa belonging to the family Asparagaceae. Some species are native to other areas such as the Caucasus. Growing from a bulb, species have linear basal leaves and a slender stalk, up to 30 cm tall, bearing clusters of typically white star-shaped flowers, often striped with green.")
			para1 = section.AddParagraph()

			' Add a paragraph to the section 
			Dim para2 As Paragraph = section.AddParagraph()
			para2.AppendText("Rosa")

			' Apply the "Heading2" style 
			para2.ApplyStyle(BuiltinStyle.Heading2)

			' Add a paragraph to the section and insert an image 
			para2 = section.AddParagraph()
			Dim picture2 As DocPicture = para2.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Rosa.jpg"))
			picture2.TextWrappingStyle = TextWrappingStyle.Square
			para2.AppendText("A rose is a woody perennial flowering plant of the genus Rosa, in the family Rosaceae, or the flower it bears. There are over a hundred species and thousands of cultivars. They form a group of plants that can be erect shrubs, climbing or trailing with stems that are often armed with sharp prickles. Flowers vary in size and shape and are usually large and showy, in colours ranging from white through yellows and reds. Most species are native to Asia, with smaller numbers native to Europe, North America, and northwestern Africa. Species, cultivars and hybrids are all widely grown for their beauty and often are fragrant. Roses have acquired cultural significance in many societies. Rose plants range in size from compact, miniature roses, to climbers that can reach seven meters in height. Different species hybridize easily, and this has been used in the development of the wide range of garden roses.")
			section.AddParagraph()

			' Add a paragraph to the section 
			Dim para3 As Paragraph = section.AddParagraph()
			para3.AppendText("Hyacinth")

			' Apply the "Heading3" style 
			para3.ApplyStyle(BuiltinStyle.Heading3)

			' Add a paragraph to the section and insert an image 
			para3 = section.AddParagraph()
			Dim picture3 As DocPicture = para3.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\hyacinths.JPG"))
			picture3.TextWrappingStyle = TextWrappingStyle.Tight
			para3.AppendText("Hyacinthus is a small genus of bulbous, fragrant flowering plants in the family Asparagaceae, subfamily Scilloideae. These are commonly called hyacinths. The genus is native to the eastern Mediterranean (from the south of Turkey through to northern Israel).")
			para3 = section.AddParagraph()
			para3.AppendText("Several species of Brodiea, Scilla, and other plants that were formerly classified in the lily family and have flower clusters borne along the stalk also have common names with the word ""hyacinth"" in them. Hyacinths should also not be confused with the genus Muscari, which are commonly known as grape hyacinths.")

			' Update the table of contents
			doc.UpdateTableOfContents()

			' Save the document to a file
			doc.SaveToFile("result.docx", FileFormat.Docx)

			' Dispose of the document object
			doc.Dispose()
			
			'Launch the Word file
			FileViewer("result.docx")
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		Protected Overrides Overloads Sub Dispose(ByVal disposing As Boolean)
			If disposing AndAlso (components IsNot Nothing) Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(Form1))
			Me.pictureBox1 = New PictureBox()
			Me.button1 = New Button()
			Me.label1 = New Label()
			CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.SuspendLayout()
			' 
			' pictureBox1
			' 
			Me.pictureBox1.Image = My.Resources.Word
			Me.pictureBox1.Location = New Point(12, 12)
			Me.pictureBox1.Name = "pictureBox1"
			Me.pictureBox1.Size = New Size(56, 48)
			Me.pictureBox1.SizeMode = PictureBoxSizeMode.Zoom
			Me.pictureBox1.TabIndex = 0
			Me.pictureBox1.TabStop = False
			' 
			' button1
			' 
			Me.button1.Anchor = (CType((AnchorStyles.Top Or AnchorStyles.Right), AnchorStyles))
			Me.button1.BackColor = Color.Transparent
			Me.button1.FlatAppearance.BorderColor = Color.FromArgb((CInt((CByte(255)))), (CInt((CByte(192)))), (CInt((CByte(128)))))
			Me.button1.FlatAppearance.MouseDownBackColor = Color.FromArgb((CInt((CByte(255)))), (CInt((CByte(224)))), (CInt((CByte(192)))))
			Me.button1.FlatAppearance.MouseOverBackColor = Color.FromArgb((CInt((CByte(255)))), (CInt((CByte(255)))), (CInt((CByte(192)))))
			Me.button1.Image = (CType(resources.GetObject("button1.Image"), Image))
			Me.button1.ImageAlign = ContentAlignment.MiddleLeft
			Me.button1.Location = New Point(376, 80)
			Me.button1.Name = "button1"
			Me.button1.Size = New Size(96, 27)
			Me.button1.TabIndex = 63
			Me.button1.Text = "Run"
			Me.button1.UseVisualStyleBackColor = False
'			Me.button1.Click += New System.EventHandler(Me.button1_Click)
			' 
			' label1
			' 
			Me.label1.Font = New Font("Verdana", 8.25F)
			Me.label1.Location = New Point(85, 12)
			Me.label1.Name = "label1"
			Me.label1.Size = New Size(387, 65)
			Me.label1.TabIndex = 64
			Me.label1.Text = resources.GetString("label1.Text")
			' 
			' Form1
			' 
			Me.AutoScaleDimensions = New SizeF(6F, 12F)
			Me.AutoScaleMode = AutoScaleMode.Font
			Me.AutoSize = True
			Me.ClientSize = New Size(484, 122)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.button1)
			Me.Controls.Add(Me.pictureBox1)
			Me.FormBorderStyle = FormBorderStyle.FixedSingle
			Me.Name = "Form1"
			Me.StartPosition = FormStartPosition.CenterScreen
			Me.Text = "Create Table Of Content By Default"
			CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)

		End Sub

		#End Region

		Private pictureBox1 As PictureBox
		Private WithEvents button1 As Button
		Private label1 As Label

	End Class
End Namespace
