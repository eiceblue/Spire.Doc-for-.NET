using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ASCIICharactersBulletStyle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
                //Create a new document
				Document document = new Document();
				Section section = document.AddSection();

				//Create the first list style based on different ASCII characters
				ListStyle listStyle1 = new ListStyle(document, ListType.Bulleted);

				//Set the style name
				listStyle1.Name = "liststyle";

				//Set the bullet character
				listStyle1.Levels[0].BulletCharacter = "\x006e";

				//Set the font name
				listStyle1.Levels[0].CharacterFormat.FontName = "Wingdings";

				//Add the list style to the document
				document.ListStyles.Add(listStyle1);

				//Create the first list style based on different ASCII characters
				ListStyle listStyle2 = new ListStyle(document, ListType.Bulleted);

				//Set the style name
				listStyle2.Name = "liststyle2";

				//Set the bullet character
				listStyle2.Levels[0].BulletCharacter = "\x0075";

				//Set the font name
				listStyle2.Levels[0].CharacterFormat.FontName = "Wingdings";

				//Add the list style to the document
				document.ListStyles.Add(listStyle2);

				//Create the third list style based on different ASCII characters
				ListStyle listStyle3 = new ListStyle(document, ListType.Bulleted);

				//Set the style name
				listStyle3.Name = "liststyle3";

				//Set the bullet character
				listStyle3.Levels[0].BulletCharacter = "\x00b2";

				//Set the font name
				listStyle3.Levels[0].CharacterFormat.FontName = "Wingdings";

				//Add the list style to the document
				document.ListStyles.Add(listStyle3);

				//Create the forth list style based on different ASCII characters
				ListStyle listStyle4 = new ListStyle(document, ListType.Bulleted);

				//Set the style name
				listStyle4.Name = "liststyle4";

				//Set the bullet character
				listStyle4.Levels[0].BulletCharacter = "\x00d8";

				//Set the font name
				listStyle4.Levels[0].CharacterFormat.FontName = "Wingdings";

				//Add the list style to the document
				document.ListStyles.Add(listStyle4);

				//Create a paragraph
				Paragraph p1 = section.Body.AddParagraph();

				//Append text
				p1.AppendText("Spire.Doc for .NET");

				//Apply the style
				p1.ListFormat.ApplyStyle(listStyle1.Name);

				//Create a paragraph
				Paragraph p2 = section.Body.AddParagraph();

				//Append text
				p2.AppendText("Spire.Doc for .NET");

				//Apply the style
				p2.ListFormat.ApplyStyle(listStyle2.Name);

				//Create a paragraph
				Paragraph p3 = section.Body.AddParagraph();

				//Append text
				p3.AppendText("Spire.Doc for .NET");

				//Apply the style
				p3.ListFormat.ApplyStyle(listStyle3.Name);

				//Create a paragraph
				Paragraph p4 = section.Body.AddParagraph();

				//Append text
				p4.AppendText("Spire.Doc for .NET");

				//Apply the style
				p4.ListFormat.ApplyStyle(listStyle4.Name);

				//Save the document
				string output = "ASCIICharactersBulletStyle_output.docx";
				document.SaveToFile(output, FileFormat.Docx);

				//Dispose the Document
				document.Dispose();

            WordDocViewer(output);
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
