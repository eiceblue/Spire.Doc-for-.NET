Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ToPDF
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
			Dim document As New Document()

			Dim section As Section = document.AddSection()
			section.PageSetup.PageSize = PageSize.A4
			section.PageSetup.Margins.Top = 72f
			section.PageSetup.Margins.Bottom = 72f
			section.PageSetup.Margins.Left = 89.85f
			section.PageSetup.Margins.Right = 89.85f

			Dim paragraph As Paragraph = section.AddParagraph()
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left
			paragraph.AppendPicture(My.Resources.Word)

			Dim p1 As String = "Microsoft Word is a word processor designed by Microsoft. " & "It was first released in 1983 under the name Multi-Tool Word for Xenix systems. " & "Subsequent versions were later written for several other platforms including " & "IBM PCs running DOS (1983), the Apple Macintosh (1984), the AT&T Unix PC (1985), " & "Atari ST (1986), SCO UNIX, OS/2, and Microsoft Windows (1989). "
			Dim p2 As String = "Microsoft Office Word instead of merely Microsoft Word. " & "The 2010 version appears to be branded as Microsoft Word, " & "once again. The current versions are Microsoft Word 2010 for Windows and 2008 for Mac."
			section.AddParagraph().AppendText(p1).CharacterFormat.FontSize = 14
			section.AddParagraph().AppendText(p2).CharacterFormat.FontSize = 14

			'Save doc file to pdf.
			document.SaveToFile("Sample.pdf", FileFormat.PDF)

			'Launching the pdf reader to open.
			FileViewer("Sample.pdf")
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
			Me.button1.Location = New Point(349, 95)
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
			Me.label1.Size = New Size(360, 63)
			Me.label1.TabIndex = 64
			Me.label1.Text = resources.GetString("label1.Text")
			' 
			' Form1
			' 
			Me.AutoScaleDimensions = New SizeF(6F, 12F)
            Me.AutoScaleMode = Windows.Forms.AutoScaleMode.Font
			Me.AutoSize = True
			Me.ClientSize = New Size(457, 134)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.button1)
			Me.Controls.Add(Me.pictureBox1)
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
			Me.Name = "Form1"
			Me.StartPosition = FormStartPosition.CenterScreen
			Me.Text = "To PDF"
			CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)

		End Sub

		#End Region

		Private pictureBox1 As PictureBox
		Private WithEvents button1 As Button
		Private label1 As Label

	End Class
End Namespace
