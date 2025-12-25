Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.OMath

Namespace AddOMMLCode
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim OmmlCodes As List(Of String) = New List(Of String) From {
			 "<m:oMath xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:mml=""http://www.w3.org/1998/Math/MathML"" ><m:r><m:t></m:t></m:r><m:eqArr><m:e><m:r><m:t>a⊂β,b⊂β,a∩b=P</m:t></m:r></m:e><m:e><m:r><m:t>a∥</m:t></m:r><m:r><m:rPr><m:sty m:val=""p"" /></m:rPr><m:t>∂</m:t></m:r><m:r><m:t>,b∥</m:t></m:r><m:r><m:rPr><m:sty m:val=""p"" /></m:rPr><m:t>∂</m:t></m:r></m:e></m:eqArr><m:r><m:t>}</m:t></m:r><m:r><m:t>⇒β∥α</m:t></m:r></m:oMath>",
			 "<m:oMath xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:mml=""http://www.w3.org/1998/Math/MathML"" ><m:eqArr><m:e><m:r><m:t>A=</m:t></m:r><m:sSub><m:e><m:r><m:t>[</m:t></m:r><m:sSub><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sub><m:r><m:t>ij</m:t></m:r></m:sub></m:sSub><m:r><m:t>]</m:t></m:r></m:e><m:sub><m:r><m:t>m×n</m:t></m:r></m:sub></m:sSub><m:r><m:t>,B=</m:t></m:r><m:sSub><m:e><m:r><m:t>[</m:t></m:r><m:sSub><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sub><m:r><m:t>ij</m:t></m:r></m:sub></m:sSub><m:r><m:t>]</m:t></m:r></m:e><m:sub><m:r><m:t>n×s</m:t></m:r></m:sub></m:sSub></m:e><m:e><m:sSub><m:e><m:r><m:t>c</m:t></m:r></m:e><m:sub><m:r><m:t>ij</m:t></m:r></m:sub></m:sSub><m:r><m:t>=</m:t></m:r><m:nary><m:naryPr><m:chr m:val=""∑"" /><m:limLoc m:val=""undOvr"" /><m:grow m:val=""1"" /><m:subHide m:val=""off"" /><m:supHide m:val=""off"" /></m:naryPr><m:sub><m:r><m:t>k=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e /></m:nary><m:sSub><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sub><m:r><m:t>ik</m:t></m:r></m:sub></m:sSub><m:sSub><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sub><m:r><m:t>kj</m:t></m:r></m:sub></m:sSub></m:e><m:e><m:r><m:t>C=AB=</m:t></m:r><m:sSub><m:e><m:r><m:t>[</m:t></m:r><m:sSub><m:e><m:r><m:t>c</m:t></m:r></m:e><m:sub><m:r><m:t>ij</m:t></m:r></m:sub></m:sSub><m:r><m:t>]</m:t></m:r></m:e><m:sub><m:r><m:t>m×s</m:t></m:r></m:sub></m:sSub><m:r><m:t>=</m:t></m:r><m:sSub><m:e><m:r><m:t>[</m:t></m:r><m:nary><m:naryPr><m:chr m:val=""∑"" /><m:limLoc m:val=""undOvr"" /><m:grow m:val=""1"" /><m:subHide m:val=""off"" /><m:supHide m:val=""off"" /></m:naryPr><m:sub><m:r><m:t>k=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e /></m:nary><m:sSub><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sub><m:r><m:t>ik</m:t></m:r></m:sub></m:sSub><m:sSub><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sub><m:r><m:t>kj</m:t></m:r></m:sub></m:sSub><m:r><m:t>]</m:t></m:r></m:e><m:sub><m:r><m:t>m×s</m:t></m:r></m:sub></m:sSub></m:e></m:eqArr></m:oMath>",
			 "<m:oMath xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:mml=""http://www.w3.org/1998/Math/MathML""><m:r><m:rPr><m:sty m:val=""p"" /></m:rPr><m:t>sin</m:t></m:r><m:r><m:t>⁡</m:t></m:r><m:r><m:t>(</m:t></m:r><m:f><m:fPr><m:type m:val=""bar"" /></m:fPr><m:num><m:r><m:t>π</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f><m:r><m:t>−α)</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:rPr><m:sty m:val=""p"" /></m:rPr><m:t>cos</m:t></m:r><m:r><m:t>⁡α</m:t></m:r></m:oMath>",
			 "<m:oMath xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:mml=""http://www.w3.org/1998/Math/MathML""><m:eqArr><m:e><m:r><m:t>S=</m:t></m:r><m:r><m:t>(</m:t></m:r><m:f><m:fPr><m:type m:val=""noBar"" /></m:fPr><m:num><m:r><m:t>N</m:t></m:r></m:num><m:den><m:r><m:t>n</m:t></m:r></m:den></m:f><m:r><m:t>)</m:t></m:r><m:r><m:t>,</m:t></m:r><m:sSub><m:e><m:r><m:t>A</m:t></m:r></m:e><m:sub><m:r><m:t>k</m:t></m:r></m:sub></m:sSub><m:r><m:t>=</m:t></m:r><m:r><m:t>(</m:t></m:r><m:f><m:fPr><m:type m:val=""noBar"" /></m:fPr><m:num><m:r><m:t>M</m:t></m:r></m:num><m:den><m:r><m:t>k</m:t></m:r></m:den></m:f><m:r><m:t>)</m:t></m:r><m:r><m:t>⋅</m:t></m:r><m:r><m:t>(</m:t></m:r><m:f><m:fPr><m:type m:val=""noBar"" /></m:fPr><m:num><m:r><m:t>N−M</m:t></m:r></m:num><m:den><m:r><m:t>n−k</m:t></m:r></m:den></m:f><m:r><m:t>)</m:t></m:r></m:e><m:e><m:r><m:t>P</m:t></m:r><m:r><m:t>(</m:t></m:r><m:sSub><m:e><m:r><m:t>A</m:t></m:r></m:e><m:sub><m:r><m:t>k</m:t></m:r></m:sub></m:sSub><m:r><m:t>)</m:t></m:r><m:r><m:t>=</m:t></m:r><m:f><m:fPr><m:type m:val=""bar"" /></m:fPr><m:num><m:r><m:t>(</m:t></m:r><m:f><m:fPr><m:type m:val=""noBar"" /></m:fPr><m:num><m:r><m:t>M</m:t></m:r></m:num><m:den><m:r><m:t>k</m:t></m:r></m:den></m:f><m:r><m:t>)</m:t></m:r><m:r><m:t>⋅</m:t></m:r><m:r><m:t>(</m:t></m:r><m:f><m:fPr><m:type m:val=""noBar"" /></m:fPr><m:num><m:r><m:t>N−M</m:t></m:r></m:num><m:den><m:r><m:t>n−k</m:t></m:r></m:den></m:f><m:r><m:t>)</m:t></m:r></m:num><m:den><m:r><m:t>(</m:t></m:r><m:f><m:fPr><m:type m:val=""noBar"" /></m:fPr><m:num><m:r><m:t>N</m:t></m:r></m:num><m:den><m:r><m:t>n</m:t></m:r></m:den></m:f><m:r><m:t>)</m:t></m:r></m:den></m:f></m:e></m:eqArr></m:oMath>",
			 "<m:oMath xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:mml=""http://www.w3.org/1998/Math/MathML"" ><m:r><m:t>(1+x</m:t></m:r><m:sSup><m:e><m:r><m:t>)</m:t></m:r></m:e><m:sup><m:r><m:t>n</m:t></m:r></m:sup></m:sSup><m:r><m:t>=1+</m:t></m:r><m:f><m:fPr><m:type m:val=""bar"" /></m:fPr><m:num><m:r><m:t>nx</m:t></m:r></m:num><m:den><m:r><m:t>1!</m:t></m:r></m:den></m:f><m:r><m:t>+</m:t></m:r><m:f><m:fPr><m:type m:val=""bar"" /></m:fPr><m:num><m:r><m:t>n(n−1)</m:t></m:r><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:num><m:den><m:r><m:t>2!</m:t></m:r></m:den></m:f><m:r><m:t>+⋯</m:t></m:r></m:oMath>"
			}

			' Create a new Document instance
			Dim document As Document = New Document()

			' Add a new section to the document
			Dim section As Section = document.AddSection()

			' Iterate through each OMML code string in the OmmlCodes array
			For Each ommlCode As String In OmmlCodes
				' Create a new OfficeMath object to represent the equation
				Dim officeMath As OfficeMath = New OfficeMath(document)

				' Set the font size for the equation
				officeMath.CharacterFormat.FontSize = 14.0F

				' Load the Office Math Markup Language (OMML) code into the object
				officeMath.FromOMMLCode(ommlCode)

				' Add a new paragraph to the section, then add the OfficeMath object as a child
				section.AddParagraph().ChildObjects.Add(officeMath)

				' Add an empty paragraph after the equation for spacing
				section.AddParagraph()
			Next

			' Define the output file name
			Dim result As String = "AddOMMLEquation.docx"

			' Save the document to a DOCX file (compatible with Word 2013)
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object to free resources
			document.Dispose()

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
