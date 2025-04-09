﻿Imports Spire.Doc
Imports System.Drawing.Printing

Namespace CustomPaperSize
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim input As String = "..\..\..\..\..\..\Data\Sample.docx"

			' Create a new instance of Document
			Dim doc As New Document()

			' Load the Word document from the specified input file
			doc.LoadFromFile(input)

			' Get the PrintDocument associated with the document
			Dim printDoc As PrintDocument = doc.PrintDocument

			' Set the paper size of the default page settings to a custom size
			printDoc.DefaultPageSettings.PaperSize = New PaperSize("custom", 900, 800)

			' Print the document
			printDoc.Print()

			' Dispose of the document object when finished using it
			doc.Dispose()
		End Sub
	End Class
End Namespace
