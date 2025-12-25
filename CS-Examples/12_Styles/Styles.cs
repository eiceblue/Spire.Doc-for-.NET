using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Interface;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Styles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize a document
			Document document = new Document();

			// Add a section
			Section sec = document.AddSection();

			// Add default title style to document and modify
			Style titleStyle = document.AddStyle(BuiltinStyle.Title);

			//judge if it is Paragraph Style and then set paragraph format
			if (titleStyle is ParagraphStyle)
			{
				ParagraphStyle ps = titleStyle as ParagraphStyle;

                //Set the font and font size
                ps.CharacterFormat.Font = new System.Drawing.Font("cambria", 28);

                //Set the text color
                ps.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136);

                //Set the BorderType
                ps.ParagraphFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single;

				//Set the color
				ps.ParagraphFormat.Borders.Bottom.Color = Color.FromArgb(42, 123, 136);

				//Set the line width
				ps.ParagraphFormat.Borders.Bottom.LineWidth = 1.5f;

				//Set the horizontal alignment style
				ps.ParagraphFormat.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
            }

            // Add the normal text style
            Style normalStyle = document.AddStyle(BuiltinStyle.Normal);

            if (normalStyle is ParagraphStyle)
            {
                ParagraphStyle ps = normalStyle as ParagraphStyle;
                ps.CharacterFormat.Font = new System.Drawing.Font("cambria", 11);
            }


            // Add default heading1 style
            Style heading1Style = document.AddStyle(BuiltinStyle.Heading1);
            if (heading1Style is ParagraphStyle)
            {
                ParagraphStyle ps = heading1Style as ParagraphStyle;
                ps.CharacterFormat.Font = new System.Drawing.Font("cambria", 14);
                ps.CharacterFormat.Bold = true;
                ps.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136);

            }

            // Add default heading2 style
            Style heading2Style = document.AddStyle(BuiltinStyle.Heading2);

            if (heading2Style is ParagraphStyle)
            {
                ParagraphStyle ps = heading2Style as ParagraphStyle;
                ps.CharacterFormat.Font = new System.Drawing.Font("cambria", 12);
                ps.CharacterFormat.Bold = true;
            }

            // Create a bulleted list style for itemized content
            ListStyle bulletList = document.Styles.Add(ListType.Bulleted, "bulletList");

            if (bulletList != null && bulletList is ICharacterStyle)
            {
                ICharacterStyle style = bulletList as ICharacterStyle;
                style.CharacterFormat.Font = new System.Drawing.Font("cambria", 12);

            }

			// Apply styles
			Paragraph paragraph = sec.AddParagraph();
			paragraph.AppendText("Your Name");
			paragraph.ApplyStyle(BuiltinStyle.Title);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Address, City, ST ZIP Code | Telephone | Email");
			paragraph.ApplyStyle(BuiltinStyle.Normal);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Objective");
			paragraph.ApplyStyle(BuiltinStyle.Heading1);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("To get started right away, just click any placeholder text (such as this) and start typing to replace it with your own.");
			paragraph.ApplyStyle(BuiltinStyle.Normal);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Education");
			paragraph.ApplyStyle(BuiltinStyle.Heading1);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("DEGREE | DATE EARNED | SCHOOL");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Major:Text");
			paragraph.ListFormat.ApplyStyle(bulletList);
			paragraph = sec.AddParagraph();
			paragraph.AppendText("Minor:Text");
			paragraph.ListFormat.ApplyStyle(bulletList);
			paragraph = sec.AddParagraph();
			paragraph.AppendText("Related coursework:Text");
			paragraph.ListFormat.ApplyStyle(bulletList);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Skills & Abilities");
			paragraph.ApplyStyle(BuiltinStyle.Heading1);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("MANAGEMENT");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Think a document that looks this good has to be difficult to format? Think again! To easily apply any text formatting you see in this document with just a click, on the Home tab of the ribbon, check out Styles.");
			paragraph.ListFormat.ApplyStyle(bulletList);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("COMMUNICATION");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("You delivered that big presentation to rave reviews. Don¡¯t be shy about it now! This is the place to show how well you work and play with others.");
			paragraph.ListFormat.ApplyStyle(bulletList);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("LEADERSHIP");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Are you president of your fraternity, head of the condo board, or a team lead for your favorite charity? You¡¯re a natural leader¡ªtell it like it is!");
			paragraph.ListFormat.ApplyStyle(bulletList);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Experience");
			paragraph.ApplyStyle(BuiltinStyle.Heading1);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("JOB TITLE | COMPANY | DATES FROM - TO");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("This is the place for a brief summary of your key responsibilities and most stellar accomplishments.");
			paragraph.ListFormat.ApplyStyle(bulletList);

            // Save the document to a DOCX file
            string filePath = "style.docx";
			document.SaveToFile(filePath, FileFormat.Docx);

            // Dispose of the document object and open the created document in MS Word
            document.Dispose();
			WordDocViewer(filePath);
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
