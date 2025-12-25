using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Documents;
using System;
using System.Windows.Forms;
using static System.Collections.Specialized.BitVector32;

namespace Text
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a Word document
            Document document = new Document();

            //Add a section
            Spire.Doc.Section sec = document.AddSection();

            //Add paragraph and apply style
            Spire.Doc.Documents.Paragraph paragraph = sec.AddParagraph();
            paragraph.AppendText("Lists");
            paragraph.ApplyStyle(BuiltinStyle.Title);
            paragraph = sec.AddParagraph();
            paragraph.AppendText("Numbered List:").CharacterFormat.Bold = true;

            //Create list style
            ListStyle listStyle = document.Styles.Add(ListType.Numbered, "numberList");
            ListLevelCollection Levels = listStyle.ListRef.Levels;
            Levels[1].NumberPrefix = "\x0000.";
            Levels[1].PatternType = ListPatternType.Arabic;
            Levels[2].NumberPrefix = "\x0000.\x0001.";
            Levels[2].PatternType = ListPatternType.Arabic;

            ListStyle bulletList = document.Styles.Add(ListType.Bulleted, "bulletList");
            //Add paragraph and apply the list style
            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 1");
            paragraph.ListFormat.ApplyStyle(listStyle);

            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 2");
            paragraph.ListFormat.ApplyStyle(listStyle);

            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 2.1");
            paragraph.ListFormat.ApplyStyle(listStyle);
            paragraph.ListFormat.ListLevelNumber = 1;

            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 2.2");
            paragraph.ListFormat.ApplyStyle(listStyle);
            paragraph.ListFormat.ListLevelNumber = 1;

            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 2.2.1");
            paragraph.ListFormat.ApplyStyle(listStyle);
            paragraph.ListFormat.ListLevelNumber = 2;
            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 2.2.2");
            paragraph.ListFormat.ApplyStyle(listStyle);
            paragraph.ListFormat.ListLevelNumber = 2;
            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 2.2.3");
            paragraph.ListFormat.ApplyStyle(listStyle);
            paragraph.ListFormat.ListLevelNumber = 2;

            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 2.3");
            paragraph.ListFormat.ApplyStyle(listStyle);
            paragraph.ListFormat.ListLevelNumber = 1;

            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 3");
            paragraph.ListFormat.ApplyStyle(listStyle);

            paragraph = sec.AddParagraph();
            paragraph.AppendText("Bulleted List:").CharacterFormat.Bold = true;

            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 1");
            paragraph.ListFormat.ApplyStyle(bulletList);
            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 2");
            paragraph.ListFormat.ApplyStyle(bulletList);

            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 2.1");
            paragraph.ListFormat.ApplyStyle(bulletList);
            paragraph.ListFormat.ListLevelNumber = 1;
            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 2.2");
            paragraph.ListFormat.ApplyStyle(bulletList);
            paragraph.ListFormat.ListLevelNumber = 1;
            paragraph = sec.AddParagraph();
            paragraph.AppendText("List Item 3");
            paragraph.ListFormat.ApplyStyle(bulletList);



            //Save doc file.
            document.SaveToFile("lists-out.docx", FileFormat.Docx);
            document.Close();

           //Launching the MS Word file.
           WordDocViewer("lists-out.docx");


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
