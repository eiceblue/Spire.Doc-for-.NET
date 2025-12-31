using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CreateFormField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
			// Create a new Document object
			Document document = new Document();

			// Add a section to the document
			Section section = document.AddSection();

			// Set page settings for the section
			SetPage(section);

			// Insert header and footer into the section
			InsertHeaderAndFooter(section);

			// Add a title to the section
			AddTitle(section);

			// Add a form to the section
			AddForm(section);

			// Save the document to a file with the specified format
			document.SaveToFile("Sample.doc", FileFormat.Doc);

			// Dispose the document object
			document.Dispose();

            //Launch the Word file.
            WordDocViewer("Sample.doc");


        }

		private void SetPage(Section section)
		{
			// Set the page size of the section to A4
			section.PageSetup.PageSize = PageSize.A4;

			// Set the top, bottom, left, and right margins of the section
			section.PageSetup.Margins.Top = 72f;
			section.PageSetup.Margins.Bottom = 72f;
			section.PageSetup.Margins.Left = 89.85f;
			section.PageSetup.Margins.Right = 89.85f;
		}

		private void InsertHeaderAndFooter(Section section)
		{
			// Add a paragraph to the header and insert a picture
			Paragraph headerParagraph = section.HeadersFooters.Header.AddParagraph();
			DocPicture headerPicture = headerParagraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Header.png"));
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
             DocPicture headerPicture = headerParagraph.AppendPicture(@"..\..\..\..\..\..\Data\Header.png");
            */
            // Add text to the header paragraph with specified font settings
            TextRange text = headerParagraph.AppendText("Demo of Spire.Doc");
			text.CharacterFormat.FontName = "Arial";
			text.CharacterFormat.FontSize = 10;
			text.CharacterFormat.Italic = true;
			headerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

			// Set border settings for the bottom border of the header paragraph
			headerParagraph.Format.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single;
			headerParagraph.Format.Borders.Bottom.Space = 0.05F;

			// Set wrapping style and alignment for the header picture
			headerPicture.TextWrappingStyle = TextWrappingStyle.Behind;
			headerPicture.HorizontalOrigin = HorizontalOrigin.Page;
			headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
			headerPicture.VerticalOrigin = VerticalOrigin.Page;
			headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top;

			// Add a paragraph to the footer and insert a picture
			Paragraph footerParagraph = section.HeadersFooters.Footer.AddParagraph();
			DocPicture footerPicture = footerParagraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Footer.png"));
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
             DocPicture footerPicture = footerParagraph.AppendPicture(@"..\..\..\..\..\..\Data\Footer.png");
            */
            // Set wrapping style and alignment for the footer picture
            footerPicture.TextWrappingStyle = TextWrappingStyle.Behind;
			footerPicture.HorizontalOrigin = HorizontalOrigin.Page;
			footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
			footerPicture.VerticalOrigin = VerticalOrigin.Page;
			footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom;

			// Append field codes for page number and number of pages to the footer paragraph
			footerParagraph.AppendField("page number", FieldType.FieldPage);
			footerParagraph.AppendText(" of ");
			footerParagraph.AppendField("number of pages", FieldType.FieldNumPages);
			footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

			// Set border settings for the top border of the footer paragraph
			footerParagraph.Format.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;
			footerParagraph.Format.Borders.Top.Space = 0.05F;
		}

		private void AddTitle(Section section)
		{
			// Add a paragraph for the title
			Paragraph title = section.AddParagraph();

			// Append the title text with specified font settings
			TextRange titleText = title.AppendText("Create Your Account");
			titleText.CharacterFormat.FontSize = 18;
			titleText.CharacterFormat.FontName = "Arial";
			titleText.CharacterFormat.TextColor = Color.FromArgb(0x00, 0x71, 0xb6);

			// Set the horizontal alignment of the title paragraph to center
			title.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			// Set the spacing after the title paragraph
			title.Format.AfterSpacing = 8;
		}

		private void AddForm(Section section)
		{
			// Create a paragraph style for description texts
			ParagraphStyle descriptionStyle = new ParagraphStyle(section.Document);
			descriptionStyle.Name = "description";
			descriptionStyle.CharacterFormat.FontSize = 12;
			descriptionStyle.CharacterFormat.FontName = "Arial";
			descriptionStyle.CharacterFormat.TextColor = Color.FromArgb(0x00, 0x45, 0x8e);
			section.Document.Styles.Add(descriptionStyle);

			// Add the first description paragraph
			Paragraph p1 = section.AddParagraph();
			String text1 = "So that we can verify your identity and find your information, "
				+ "please provide us with the following information. "
				+ "This information will be used to create your online account. "
				+ "Your information is not public, shared in any way, or displayed on this site";
			p1.AppendText(text1);
			p1.ApplyStyle(descriptionStyle.Name);

			// Add the second description paragraph
			Paragraph p2 = section.AddParagraph();
			String text2 = "You must provide a real email address to which we will send your password.";
			p2.AppendText(text2);
			p2.ApplyStyle(descriptionStyle.Name);
			p2.Format.AfterSpacing = 8;

			// Create a paragraph style for form field group labels
			ParagraphStyle formFieldGroupLabelStyle = new ParagraphStyle(section.Document);
			formFieldGroupLabelStyle.Name = "formFieldGroupLabel";
			formFieldGroupLabelStyle.ApplyBaseStyle("description");
			formFieldGroupLabelStyle.CharacterFormat.Bold = true;
			formFieldGroupLabelStyle.CharacterFormat.TextColor = Color.White;
			section.Document.Styles.Add(formFieldGroupLabelStyle);

			// Create a paragraph style for form field labels
			ParagraphStyle formFieldLabelStyle = new ParagraphStyle(section.Document);
			formFieldLabelStyle.Name = "formFieldLabel";
			formFieldLabelStyle.ApplyBaseStyle("description");
			formFieldLabelStyle.ParagraphFormat.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
			section.Document.Styles.Add(formFieldLabelStyle);

			// Add a table to the section for the form fields
			Table table = section.AddTable();
			// Set the number of columns
			table.DefaultColumnsNumber = 2; 
			// Set the default row height
			table.DefaultRowHeight = 20; 

			// Read the XML file containing the form structure
			using (Stream stream = File.OpenRead(@"..\..\..\..\..\..\Data\Form.xml"))
			{
				XPathDocument xpathDoc = new XPathDocument(stream);
				XPathNodeIterator sectionNodes = xpathDoc.CreateNavigator().Select("/form/section");

				// Iterate over each section node in the XML file
				foreach (XPathNavigator node in sectionNodes)
				{
					// Add a row for the form field group label
					TableRow row = table.AddRow(false);
					row.Cells[0].CellFormat.Shading.BackgroundPatternColor= Color.FromArgb(0x00, 0x71, 0xb6);
					row.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

					// Add the form field group label text to the cell
					Paragraph cellParagraph = row.Cells[0].AddParagraph();
					cellParagraph.AppendText(node.GetAttribute("name", ""));
					cellParagraph.ApplyStyle(formFieldGroupLabelStyle.Name);

					// Iterate over each field node within the section node
					XPathNodeIterator fieldNodes = node.Select("field");
					foreach (XPathNavigator fieldNode in fieldNodes)
					{
						// Add a row for the form field label and input field
						TableRow fieldRow = table.AddRow(false);

						// Set vertical alignment for the cells in the field row
						fieldRow.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
						fieldRow.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

						// Add the form field label to the first cell in the row
						Paragraph labelParagraph = fieldRow.Cells[0].AddParagraph();
						labelParagraph.AppendText(fieldNode.GetAttribute("label", ""));
						labelParagraph.ApplyStyle(formFieldLabelStyle.Name);

						// Add the input field paragraph to the second cell in the row
						Paragraph fieldParagraph = fieldRow.Cells[1].AddParagraph();
						String fieldId = fieldNode.GetAttribute("id", "");
						switch (fieldNode.GetAttribute("type", ""))
						{
							case "text":
								// Add a text form input field
								TextFormField field = fieldParagraph.AppendField(fieldId, FieldType.FieldFormTextInput) as TextFormField;
								field.DefaultText = "";
								field.Text = "";
								break;

                            case "list":
                                // Add a dropdown list form field
                                DropDownFormField list
                                    = fieldParagraph.AppendField(fieldId, FieldType.FieldFormDropDown) as DropDownFormField;

                              
                                XPathNodeIterator itemNodes = fieldNode.Select("item");
                                foreach (XPathNavigator itemNode in itemNodes)
                                {
                                    list.DropDownItems.Add(itemNode.SelectSingleNode("text()").Value);
                                }
                                break;

                            case "checkbox":
                                // Add a checkbox form field
                                fieldParagraph.AppendField(fieldId, FieldType.FieldFormCheckBox);
                                break;
                        }
                    }

                  
                    table.ApplyHorizontalMerge(row.GetRowIndex(), 0, 1);
                }
            }
        }

        private void WordDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}
