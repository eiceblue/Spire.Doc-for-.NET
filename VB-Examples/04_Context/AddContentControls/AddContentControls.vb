Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddContentControls
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			'Creat a new word document
			Dim document As New Document()
			Dim section As Section = document.AddSection()
			Dim paragraph As Paragraph = section.AddParagraph()

			'Add Combo Box Content Control
			Dim sd As New StructureDocumentTagInline(document)
			paragraph.ChildObjects.Add(sd)
			sd.SDTProperties.SDTType = SdtType.ComboBox
			Dim cb As New SdtComboBox()
			cb.ListItems.Add(New SdtListItem("Cat"))
			cb.ListItems.Add(New SdtListItem("Dog"))
			sd.SDTProperties.ControlProperties = cb
			Dim rt As New TextRange(document)
			rt.Text = cb.ListItems(0).DisplayText
			sd.SDTContent.ChildObjects.Add(rt)

			section.AddParagraph()

			'Add Text Content Control
			paragraph = section.AddParagraph()
			sd = New StructureDocumentTagInline(document)
			paragraph.ChildObjects.Add(sd)
			sd.SDTProperties.SDTType = SdtType.Text
			Dim text As New SdtText(True)
			text.IsMultiline = True
			sd.SDTProperties.ControlProperties = text
			rt = New TextRange(document)
			rt.Text = "Text"
			sd.SDTContent.ChildObjects.Add(rt)

			section.AddParagraph()
			'Add Picture Content Control
			paragraph = section.AddParagraph()
			sd = New StructureDocumentTagInline(document)
			paragraph.ChildObjects.Add(sd)
            Dim pic As New DocPicture(document)
            pic.Width = 10
            pic.Height = 10
            pic.LoadImage(Image.FromFile("..\..\..\..\..\..\Data\log.png"))
			sd.SDTContent.ChildObjects.Add(pic)

			section.AddParagraph()
			'Add Date Picker Content Control
			paragraph = section.AddParagraph()
			sd = New StructureDocumentTagInline(document)
			paragraph.ChildObjects.Add(sd)
			sd.SDTProperties.SDTType = SdtType.DatePicker
			Dim [date] As New SdtDate()
			[date].CalendarType = CalendarType.Default
			[date].DateFormat = "yyyy.MM.dd"
			[date].FullDate = Date.Now
			sd.SDTProperties.ControlProperties = [date]
			rt = New TextRange(document)
			rt.Text = "1990.02.08"
			sd.SDTContent.ChildObjects.Add(rt)

			section.AddParagraph()
			'Add Drop-Down List Content Control
			paragraph = section.AddParagraph()
			sd = New StructureDocumentTagInline(document)
			paragraph.ChildObjects.Add(sd)
			sd.SDTProperties.SDTType = SdtType.DropDownList
			Dim sddl As New SdtDropDownList()
			sddl.ListItems.Add(New SdtListItem("Harry"))
			sddl.ListItems.Add(New SdtListItem("Jerry"))
			sd.SDTProperties.ControlProperties = sddl
			rt = New TextRange(document)
			rt.Text = sddl.ListItems(0).DisplayText
			sd.SDTContent.ChildObjects.Add(rt)

			'Save and launch the file
			Dim resultfile As String = "sample.docx"
			document.SaveToFile(resultfile, FileFormat.Docx)

			FileViewer(resultfile)
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
			Me.Text = "Add Content Controls"
			CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)

		End Sub

		#End Region

		Private pictureBox1 As PictureBox
		Private WithEvents button1 As Button
		Private label1 As Label

	End Class
End Namespace
