using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

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
            //Initialize a document
			Document document = new Document();

			//Add a section
			Section sec = document.AddSection();

			//Add default title style to document and modify
			Style titleStyle = document.AddStyle(BuiltinStyle.Title);

			//Set the font and font size
			titleStyle.CharacterFormat.Font = new System.Drawing.Font("cambria", 28);
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            titleStyle.CharacterFormat.FontName= "cambria";
            titleStyle.CharacterFormat.FontSize = 28;
            */

            //Set the text color
            titleStyle.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136);

			//judge if it is Paragraph Style and then set paragraph format
			if (titleStyle is ParagraphStyle)
			{
				ParagraphStyle ps = titleStyle as ParagraphStyle;

				//Set the BorderType
				ps.ParagraphFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single;

				//Set the color
				ps.ParagraphFormat.Borders.Bottom.Color = Color.FromArgb(42, 123, 136);

				//Set the line width
				ps.ParagraphFormat.Borders.Bottom.LineWidth = 1.5f;

				//Set the horizontal alignment style
				ps.ParagraphFormat.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
			}

			//Add default normal style and modify
			Style normalStyle = document.AddStyle(BuiltinStyle.Normal);
			normalStyle.CharacterFormat.Font = new System.Drawing.Font("cambria", 11);
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            normalStyle.CharacterFormat.FontName = "cambria";
            normalStyle.CharacterFormat.FontSize = 11;
            */

            //Add default heading1 style
            Style heading1Style = document.AddStyle(BuiltinStyle.Heading1);
			heading1Style.CharacterFormat.Font = new System.Drawing.Font("cambria", 14);
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            heading1Style.CharacterFormat.FontName = "cambria";
            heading1Style.CharacterFormat.FontSize = 14;
            */
            heading1Style.CharacterFormat.Bold = true;
			heading1Style.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136);

			//Add default heading2 style
			Style heading2Style = document.AddStyle(BuiltinStyle.Heading2);
			heading2Style.CharacterFormat.Font = new System.Drawing.Font("cambria", 12);
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            heading2Style.CharacterFormat.FontName = "cambria";
            heading2Style.CharacterFormat.FontSize = 12;
            */
            heading2Style.CharacterFormat.Bold = true;

			//Create a bulletList
			ListStyle bulletList = new ListStyle(document, ListType.Bulleted);
			bulletList.CharacterFormat.Font = new System.Drawing.Font("cambria", 12);
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            bulletList.CharacterFormat.FontName = "cambria";
            bulletList.CharacterFormat.FontSize = 12;
            */

            //Set the bulletList name
            bulletList.Name = "bulletList";

			//Add the style
			document.ListStyles.Add(bulletList);


			//Apply the Title style
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
			paragraph.ListFormat.ApplyStyle("bulletList");
			paragraph = sec.AddParagraph();
			paragraph.AppendText("Minor:Text");
			paragraph.ListFormat.ApplyStyle("bulletList");
			paragraph = sec.AddParagraph();
			paragraph.AppendText("Related coursework:Text");
			paragraph.ListFormat.ApplyStyle("bulletList");

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Skills & Abilities");
			paragraph.ApplyStyle(BuiltinStyle.Heading1);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("MANAGEMENT");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Think a document that looks this good has to be difficult to format? Think again! To easily apply any text formatting you see in this document with just a click, on the Home tab of the ribbon, check out Styles.");
			paragraph.ListFormat.ApplyStyle("bulletList");

			paragraph = sec.AddParagraph();
			paragraph.AppendText("COMMUNICATION");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("You delivered that big presentation to rave reviews. Don¡¯t be shy about it now! This is the place to show how well you work and play with others.");
			paragraph.ListFormat.ApplyStyle("bulletList");

			paragraph = sec.AddParagraph();
			paragraph.AppendText("LEADERSHIP");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Are you president of your fraternity, head of the condo board, or a team lead for your favorite charity? You¡¯re a natural leader¡ªtell it like it is!");
			paragraph.ListFormat.ApplyStyle("bulletList");

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Experience");
			paragraph.ApplyStyle(BuiltinStyle.Heading1);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("JOB TITLE | COMPANY | DATES FROM - TO");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("This is the place for a brief summary of your key responsibilities and most stellar accomplishments.");
			paragraph.ListFormat.ApplyStyle("bulletList");

			//Save to docx file.
			string filePath = "Sample.docx";
			document.SaveToFile(filePath, FileFormat.Docx);

			//Dispose the document
			document.Dispose();
			
            //Launching the MS Word file.
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
