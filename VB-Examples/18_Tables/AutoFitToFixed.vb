Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AutoFitToFixed
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a document
			Dim document As New Document()

			'Load file
			document.LoadFromFile("..\..\..\..\..\..\Data\TableSample.docx")

			'Get the first section
			Dim section As Section = document.Sections(0)

			'Get the first table
			Dim table As Table = TryCast(section.Tables(0), Table)

			'The table is set to a fixed size
			table.AutoFit(AutoFitBehaviorType.FixedColumnWidths)

			'Save to file
			document.SaveToFile("result.docx")

			'Dispose the document
			document.Dispose()
			
			'Launch the document
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
			Me.Text = "Auto Fit To Fixed"
			CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)

		End Sub

		#End Region

		Private pictureBox1 As PictureBox
		Private WithEvents button1 As Button
		Private label1 As Label

	End Class
End Namespace
