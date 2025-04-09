Imports System.ComponentModel
Imports System.Net
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace HandleUrlWhileConvertHtmlToPDF
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Subscribe to the HtmlUrlLoadEvent to handle external resource loading.
			AddHandler document.HtmlUrlLoadEvent, AddressOf MyDownloadEvent

			' Load an HTML file into the document. The file path and validation type are specified.
			document.LoadFromFile("..\..\..\..\..\..\..\Data\Template_HtmlFile3.html", FileFormat.Html, XHTMLValidationType.None)

			' Save the loaded HTML content as a PDF file.
			document.SaveToFile("HtmlFileToPDF.pdf", FileFormat.PDF)

			' Dispose the document object to release resources.
			document.Dispose()
			'Launching the pdf reader to open.
			FileViewer("HtmlFileToPDF.pdf")
		End Sub

		Private Shared Sub MyDownloadEvent(ByVal sender As Object, ByVal args As Document.HtmlUrlLoadEventArgs)
			' Use WebClient to download external resources (e.g., images, CSS files) from URLs referenced in the HTML.
			Using webClient As New WebClient()
				' Use the default credentials for authentication.
				webClient.Credentials = CredentialCache.DefaultCredentials

				' Set a custom user-agent header to mimic a web browser during resource download.
				webClient.Headers.Set("user-agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0")

				' Configure supported security protocols for SSL/TLS connections.
				' This ensures compatibility with different server configurations.
				' SystemDefault = 0, Ssl3 = 48, Tls = 192, Tls11 = 768, Tls12 = 3072, Tls13 = 12288
				ServicePointManager.SecurityProtocol = CType(0, SecurityProtocolType) Or CType(12288, SecurityProtocolType) Or CType(3072, SecurityProtocolType) Or CType(768, SecurityProtocolType) Or CType(192, SecurityProtocolType) Or CType(48, SecurityProtocolType) ' Ssl3 -  Tls -  Tls11 -  Tls12 -  Tls13

				' Download the resource data from the provided URL.
				Dim webData() As Byte = webClient.DownloadData(args.Url)

				' Set the downloaded data into the event arguments for further processing.
				args.DataBytes = webData
			End Using
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
			Me.Text = "Handle Url While Convert Html To PDF"
			CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)

		End Sub

		#End Region

		Private pictureBox1 As PictureBox
		Private WithEvents button1 As Button
		Private label1 As Label

	End Class
End Namespace
