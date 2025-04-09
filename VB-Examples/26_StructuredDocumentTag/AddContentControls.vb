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
			' Create a new document object
			Dim document As New Document()

			' Add a section to the document
			Dim section As Section = document.AddSection()

			' Add a paragraph to the section and append text
			Dim paragraph As Paragraph = section.AddParagraph()
			Dim txtRange As TextRange = paragraph.AppendText("The following example shows how to add content controls in a Word document.")

			' Add an empty paragraph to the section
			paragraph = section.AddParagraph()

			' Add a paragraph to the section and append text indicating adding a ComboBox content control
			paragraph = section.AddParagraph()
			txtRange = paragraph.AppendText("Add Combo Box Content Control:  ")
			txtRange.CharacterFormat.Italic = True
			Dim sd As New StructureDocumentTagInline(document)
			paragraph.ChildObjects.Add(sd)
			sd.SDTProperties.SDTType = SdtType.ComboBox
			Dim cb As New SdtComboBox()
			cb.ListItems.Add(New SdtListItem("Spire.Doc"))
			cb.ListItems.Add(New SdtListItem("Spire.XLS"))
			cb.ListItems.Add(New SdtListItem("Spire.PDF"))
			sd.SDTProperties.ControlProperties = cb
			Dim rt As New TextRange(document)
			rt.Text = cb.ListItems(0).DisplayText
			sd.SDTContent.ChildObjects.Add(rt)

			' Add an empty paragraph to the section
			section.AddParagraph()

			' Add a paragraph to the section and append text indicating adding a Text content control
			paragraph = section.AddParagraph()
			txtRange = paragraph.AppendText("Add Text Content Control:  ")
			txtRange.CharacterFormat.Italic = True
			sd = New StructureDocumentTagInline(document)
			paragraph.ChildObjects.Add(sd)
			sd.SDTProperties.SDTType = SdtType.Text
			Dim text As New SdtText(True)
			text.IsMultiline = True
			sd.SDTProperties.ControlProperties = text
			rt = New TextRange(document)
			rt.Text = "Text"
			sd.SDTContent.ChildObjects.Add(rt)

			' Add an empty paragraph to the section
			section.AddParagraph()

			' Add a paragraph to the section and append text indicating adding a Picture content control
			paragraph = section.AddParagraph()
			txtRange = paragraph.AppendText("Add Picture Content Control:  ")
			txtRange.CharacterFormat.Italic = True
			sd = New StructureDocumentTagInline(document)
			paragraph.ChildObjects.Add(sd)
			sd.SDTProperties.SDTType = SdtType.Picture
			Dim pic As New DocPicture(document)
			pic.Width = 10
			pic.Height = 10
			pic.LoadImage(Image.FromFile("..\..\..\..\..\..\Data\logo.png"))
			sd.SDTContent.ChildObjects.Add(pic)

			' Add an empty paragraph to the section
			section.AddParagraph()

			' Add a paragraph to the section and append text indicating adding a Date Picker content control
			paragraph = section.AddParagraph()
			txtRange = paragraph.AppendText("Add Date Picker Content Control:  ")
			txtRange.CharacterFormat.Italic = True
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

			' Add an empty paragraph to the section
			section.AddParagraph()

			' Add a paragraph to the section and append text indicating adding a Drop-Down List content control
			paragraph = section.AddParagraph()
			txtRange = paragraph.AppendText("Add Drop-Down List Content Control:  ")
			txtRange.CharacterFormat.Italic = True
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

			' Specify the output file name
			Dim resultfile As String = "Output.docx"

			' Save the document to a file in Docx format
			document.SaveToFile(resultfile, FileFormat.Docx)

			' Dispose the document object
			document.Dispose()

			'Launch the Word file.
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
			Me.button1.Location = New Point(376, 83)
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
