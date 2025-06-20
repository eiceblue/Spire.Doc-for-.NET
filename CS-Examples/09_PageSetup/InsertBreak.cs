using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertBreak
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class
			Document document = new Document();

			// Add a section to the document
			Section section = document.AddSection();

			// Set page settings for the section
			SetPage(section);

			// Insert a cover page in the section
			InsertCover(section);

			// Add another section to the document
			section = document.AddSection();

			// Insert a section break at the beginning of the section
			section.AddParagraph().InsertSectionBreak(SectionBreakType.NewPage);

			// Insert content into the section
			InsertContent(section);

			// Save the document to a file named "Sample.docx" in DOCX format
			document.SaveToFile("Sample.docx", FileFormat.Docx);

			// Release the resources used by the document object
			document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");
        }

        private static void SetPage(Section section)
		{
			// Set the page size of the section to A4
			section.PageSetup.PageSize = PageSize.A4;

			// Set the top margin of the section to 72 points
			section.PageSetup.Margins.Top = 72f;

			// Set the bottom margin of the section to 72 points
			section.PageSetup.Margins.Bottom = 72f;

			// Set the left margin of the section to 89.85 points
			section.PageSetup.Margins.Left = 89.85f;

			// Set the right margin of the section to 89.85 points
			section.PageSetup.Margins.Right = 89.85f;
		}

		private static void InsertCover(Section section)
		{
			// Create a paragraph style for small text
			ParagraphStyle small = new ParagraphStyle(section.Document);
			small.Name = "small";
			small.CharacterFormat.FontName = "Arial";
			small.CharacterFormat.FontSize = 9;
			small.CharacterFormat.TextColor = Color.Gray;
			section.Document.Styles.Add(small);

			// Add a paragraph with small text
			Paragraph paragraph = section.AddParagraph();
			paragraph.AppendText("The sample demonstrates how to insert section break.");
			paragraph.ApplyStyle(small.Name);

			// Add a title paragraph
			Paragraph title = section.AddParagraph();
			TextRange text = title.AppendText("Field Types Supported by Spire.Doc");
			text.CharacterFormat.FontName = "Arial";
			text.CharacterFormat.FontSize = 20;
			text.CharacterFormat.Bold = true;

			// Set formatting for the title paragraph
			title.Format.BeforeSpacing = section.PageSetup.PageSize.Height / 2 - 3 * section.PageSetup.Margins.Top;
			title.Format.AfterSpacing = 8;
			title.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

			// Add a paragraph with small text after the title
			paragraph = section.AddParagraph();
			paragraph.AppendText("e-iceblue Spire.Doc team.");
			paragraph.ApplyStyle(small.Name);
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
		}

		private static void InsertContent(Section section)
		{
			// Create a paragraph style for the list items
			ParagraphStyle list = new ParagraphStyle(section.Document);
			list.Name = "list";
			list.CharacterFormat.FontName = "Arial";
			list.CharacterFormat.FontSize = 11;
			list.ParagraphFormat.LineSpacing = 1.5F * 12F;
			list.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
			section.Document.Styles.Add(list);

			// Add a title paragraph for the field type list
			Paragraph title = section.AddParagraph();
			TextRange text = title.AppendText("Field type list:");
			title.ApplyStyle(list.Name);

			bool first = true;
			foreach (FieldType type in Enum.GetValues(typeof(FieldType)))
			{
				// Skip unsupported or invalid field types
				if (type == FieldType.FieldUnknown || type == FieldType.FieldNone || type == FieldType.FieldEmpty)
				{
					continue;
				}

				// Add a paragraph for each supported field type
				Paragraph paragraph = section.AddParagraph();
				paragraph.AppendText(String.Format("{0} is supported in Spire.Doc", type));

				// Apply numbered or continued numbering list formatting
				if (first)
				{
					paragraph.ListFormat.ApplyNumberedStyle();
					first = false;
				}
				else
				{
					paragraph.ListFormat.ContinueListNumbering();
				}
				paragraph.ApplyStyle(list.Name);
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
