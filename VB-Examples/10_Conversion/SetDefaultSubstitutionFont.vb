Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetDefaultSubstitutionFont
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            ' Create a new Document instance
            Dim document As Document = New Document()

            ' Set the default substitution font name to "Arial"
            ' This font will be used if a specified font is not available
            document.DefaultSubstitutionFontName = "Arial"

            ' Add a new section to the document
            Dim section As Section = document.AddSection()

            ' Add a new paragraph to the section
            Dim paragraph As Paragraph = section.AddParagraph()

            ' Append text to the paragraph and get a reference to the text range
            Dim textRange As TextRange = paragraph.AppendText("Welcome to evaluate Spire.Doc for .NET product.")

            ' Set the font name of the text range to "San Francisco"
            ' (This font might not be available on the system)
            textRange.CharacterFormat.FontName = "San Francisco"

            ' Set the font size of the text range to 16
            textRange.CharacterFormat.FontSize = 16

            ' Define the output file name
            Dim result As String = "SetDefaultSubstitutionFont-result.pdf"

            ' Save the document to a PDF file
            document.SaveToFile(result, FileFormat.PDF)

            ' Dispose of the Document object to release resources
            document.Dispose()

            'Launching the pdf reader to open.
            FileViewer("SetDefaultSubstitutionFont-result.pdf")
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
		Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
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
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
            Me.pictureBox1 = New System.Windows.Forms.PictureBox()
            Me.button1 = New System.Windows.Forms.Button()
            Me.label1 = New System.Windows.Forms.Label()
            CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'pictureBox1
            '
            Me.pictureBox1.Image = Global.My.Resources.Resources.Word
            Me.pictureBox1.Location = New System.Drawing.Point(12, 12)
            Me.pictureBox1.Name = "pictureBox1"
            Me.pictureBox1.Size = New System.Drawing.Size(56, 48)
            Me.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
            Me.pictureBox1.TabIndex = 0
            Me.pictureBox1.TabStop = False
            '
            'button1
            '
            Me.button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.button1.BackColor = System.Drawing.Color.Transparent
            Me.button1.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
            Me.button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
            Me.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
            Me.button1.Image = CType(resources.GetObject("button1.Image"), System.Drawing.Image)
            Me.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.button1.Location = New System.Drawing.Point(376, 80)
            Me.button1.Name = "button1"
            Me.button1.Size = New System.Drawing.Size(96, 27)
            Me.button1.TabIndex = 63
            Me.button1.Text = "Run"
            Me.button1.UseVisualStyleBackColor = False
            '
            'label1
            '
            Me.label1.Font = New System.Drawing.Font("Verdana", 8.25!)
            Me.label1.Location = New System.Drawing.Point(85, 12)
            Me.label1.Name = "label1"
            Me.label1.Size = New System.Drawing.Size(387, 65)
            Me.label1.TabIndex = 64
            Me.label1.Text = resources.GetString("label1.Text")
            '
            'Form1
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.AutoSize = True
            Me.ClientSize = New System.Drawing.Size(484, 122)
            Me.Controls.Add(Me.label1)
            Me.Controls.Add(Me.button1)
            Me.Controls.Add(Me.pictureBox1)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
            Me.Name = "Form1"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Set the default substitution font"
            CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub

#End Region

        Private pictureBox1 As PictureBox
		Private WithEvents button1 As Button
		Private label1 As Label

	End Class
End Namespace
