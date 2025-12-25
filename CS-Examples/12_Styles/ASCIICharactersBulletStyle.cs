using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

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

            //Create four list styles based on different ASCII characters
            ListStyle listStyle1 = document.Styles.Add(ListType.Bulleted, "liststyle");
            ListLevelCollection Levels = listStyle1.ListRef.Levels;

            Levels[0].BulletCharacter = "\x006e";
            Levels[0].CharacterFormat.FontName = "Wingdings";

            //ListStyle listStyle2 = new ListStyle(document, ListType.Bulleted);
            //listStyle2.Name = "liststyle2";
            ListStyle listStyle2 = document.Styles.Add(ListType.Bulleted, "liststyle2");
            ListLevelCollection Levels2 = listStyle2.ListRef.Levels;
            Levels2[0].BulletCharacter = "\x0075";
            Levels2[0].CharacterFormat.FontName = "Wingdings";

            ListStyle listStyle3 = document.Styles.Add(ListType.Bulleted, "liststyle3");
            ListLevelCollection Levels3 = listStyle3.ListRef.Levels;
            Levels3[0].BulletCharacter = "\x00b2";
            Levels3[0].CharacterFormat.FontName = "Wingdings";

            ListStyle listStyle4 = document.Styles.Add(ListType.Bulleted, "liststyle4");
            ListLevelCollection Levels4 = listStyle4.ListRef.Levels;
            Levels4[0].BulletCharacter = "\x00d8";
            Levels4[0].CharacterFormat.FontName = "Wingdings";

            //Add four paragraphs and apply list style separately
            Paragraph p1 = section.Body.AddParagraph();
            p1.AppendText("Spire.Doc for .NET");
            p1.ListFormat.ApplyStyle(listStyle1.Name);
            Paragraph p2 = section.Body.AddParagraph();
            p2.AppendText("Spire.Doc for .NET");
            p2.ListFormat.ApplyStyle(listStyle2.Name);
            Paragraph p3 = section.Body.AddParagraph();
            p3.AppendText("Spire.Doc for .NET");
            p3.ListFormat.ApplyStyle(listStyle3.Name);
            Paragraph p4 = section.Body.AddParagraph();
            p4.AppendText("Spire.Doc for .NET");
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
