Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc

Namespace ToPCL
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile("..\..\..\..\..\..\Data\ConvertedTemplate.docx")
                        
                         ' On Net4.6 and above platforms with adding the following external dependencies, you can set the UseHarfBuzzTextShaper which can better handling Thai and Tibetan characters
          	  	' external reference to:  
            		' HarfBuzzSharp >= 2.6.1.5
            		' System.Buffers >= 4.4.0
            		' System.Memory >= 4.5.3
            		' System.Numerics.Vectors >= 4.4.0
            		' System.Runtime.CompilerServices.Unsafe >= 4.5.2

            		' document.LayoutOptions.UseHarfBuzzTextShaper = True

			Dim result As String = "ToPCL.pcl"

			'Save to file
			doc.SaveToFile(result, FileFormat.PCL)

			'Dispose the document
			doc.Dispose()
			
			WordDocViewer(result)
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
