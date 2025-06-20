using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;
using Spire.Doc.Fields;

namespace RemoveHyperlinks
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
         
			// Specify the input file path for the document containing hyperlinks
			string input = @"..\..\..\..\..\..\Data\Hyperlinks.docx";

			// Create a new Document object
			Document doc = new Document();

			// Load the document from the specified file path
			doc.LoadFromFile(input);

			// Find all the hyperlinks in the document and store them in a list
			List<Field> hyperlinks = FindAllHyperlinks(doc);

			// Flatten each hyperlink, removing the hyperlink functionality but keeping the text
			for (int i = hyperlinks.Count - 1; i >= 0; i--)
			{
				FlattenHyperlinks(hyperlinks[i]);
			}

			// Specify the output file path for the modified document without hyperlinks
			string output = "RemoveHyperlinks.docx";

			// Save the modified document to the output file path in DOCX format
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the document object to free up resources
			doc.Dispose();
			
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

		   
		// Method to find all hyperlinks in the document and return them as a list
		private List<Field> FindAllHyperlinks(Document document)
		{
			List<Field> hyperlinks = new List<Field>();

			foreach (Section section in document.Sections)
			{
				foreach (DocumentObject sec in section.Body.ChildObjects)
				{
					if (sec.DocumentObjectType == DocumentObjectType.Paragraph)
					{
						foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
						{
							if (para.DocumentObjectType == DocumentObjectType.Field)
							{
								Field field = para as Field;
								if (field.Type == FieldType.FieldHyperlink)
								{
									hyperlinks.Add(field);
								}
							}
						}
					}
				}
			}
			return hyperlinks;
		}

    
		// Method to flatten a hyperlink, removing the hyperlink functionality but keeping the text
		private void FlattenHyperlinks(Field field)
		{
			// Store the indices of relevant objects for later removal
			int ownerParaIndex = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.OwnerParagraph);
			int fieldIndex = field.OwnerParagraph.ChildObjects.IndexOf(field);
			Paragraph sepOwnerPara = field.Separator.OwnerParagraph;
			int sepOwnerParaIndex = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.Separator.OwnerParagraph);
			int sepIndex = field.Separator.OwnerParagraph.ChildObjects.IndexOf(field.Separator);
			int endIndex = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End);
			int endOwnerParaIndex = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.End.OwnerParagraph);

			// Format the text between the separator and the end of the field result
			FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex);

			// Remove the end field marker
			field.End.OwnerParagraph.ChildObjects.RemoveAt(endIndex);

			// Remove the field and its associated objects in reverse order
			for (int i = sepOwnerParaIndex; i >= ownerParaIndex; i--)
			{
				if (i == sepOwnerParaIndex && i == ownerParaIndex)
				{
					// Remove objects from the same paragraph as the field
					for (int j = sepIndex; j >= fieldIndex; j--)
					{
						field.OwnerParagraph.ChildObjects.RemoveAt(j);
					}
				}
				else if (i == ownerParaIndex)
				{
					// Remove objects from the field's paragraph but after the field
					for (int j = field.OwnerParagraph.ChildObjects.Count - 1; j >= fieldIndex; j--)
					{
						field.OwnerParagraph.ChildObjects.RemoveAt(j);
					}
				}
				else if (i == sepOwnerParaIndex)
				{
					// Remove objects from the separator's paragraph
					for (int j = sepIndex; j >= 0; j--)
					{
						sepOwnerPara.ChildObjects.RemoveAt(j);
					}
				}
				else
				{
					// Remove objects from other paragraphs
					field.OwnerParagraph.OwnerTextBody.ChildObjects.RemoveAt(i);
				}
			}
		}

		// Method to format the text between the separator and the end of a field result in the document body
		private void FormatFieldResultText(Body ownerBody, int sepOwnerParaIndex, int endOwnerParaIndex, int sepIndex, int endIndex)
		{
			for (int i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++)
			{
				// Get the paragraph at the current index
				Paragraph para = ownerBody.ChildObjects[i] as Paragraph;
				
				if (i == sepOwnerParaIndex && i == endOwnerParaIndex)
				{
					// Format objects within the same paragraph as the separator and the end of the field
					for (int j = sepIndex + 1; j < endIndex; j++)
					{
						FormatText(para.ChildObjects[j] as TextRange);
					}
				}
				else if (i == sepOwnerParaIndex)
				{
					// Format objects after the separator in the separator's paragraph
					for (int j = sepIndex + 1; j < para.ChildObjects.Count; j++)
					{
						FormatText(para.ChildObjects[j] as TextRange);
					}
				}
				else if (i == endOwnerParaIndex)
				{
					// Format objects before the end of the field in the end paragraph
					for (int j = 0; j < endIndex; j++)
					{
						FormatText(para.ChildObjects[j] as TextRange);
					}
				}
				else
				{
					// Format all objects in other paragraphs
					for (int j = 0; j < para.ChildObjects.Count; j++)
					{
						FormatText(para.ChildObjects[j] as TextRange);
					}
				}
			}
		}

		// Method to format the text range by setting its color to black and removing underline style
		private void FormatText(TextRange tr)
		{
			tr.CharacterFormat.TextColor = Color.Black;
			tr.CharacterFormat.UnderlineStyle = UnderlineStyle.None;
		}
    }
}
