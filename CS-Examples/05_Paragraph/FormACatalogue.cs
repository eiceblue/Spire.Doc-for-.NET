using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Documents;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FormACatalogue
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object.
            Document document = new Document();

            //Add a new section. 
            Section section = document.AddSection();
            Spire.Doc.Documents.Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

            //Add Heading 1.
            paragraph = section.AddParagraph();
            paragraph.AppendText(BuiltinStyle.Heading1.ToString());
            paragraph.ApplyStyle(BuiltinStyle.Heading1);
            paragraph.ListFormat.ApplyNumberedStyle();

            //Add Heading 2.
            paragraph = section.AddParagraph();
            paragraph.AppendText(BuiltinStyle.Heading2.ToString());
            paragraph.ApplyStyle(BuiltinStyle.Heading2);

            //List style for Headings 2.

            ListStyle listStyle2 = document.Styles.Add(ListType.Numbered, "MyStyle2");
            ListLevelCollection Levels = listStyle2.ListRef.Levels;
            foreach (ListLevel listLev in Levels)
            {
                listLev.UsePrevLevelPattern = true;
                listLev.NumberPrefix = "1.";
            }

            paragraph.ListFormat.ApplyStyle(listStyle2.Name);

            //Add list style 3.

            ListStyle listStyle3 = document.Styles.Add(ListType.Numbered, "MyStyle3");
            ListLevelCollection Levels1 = listStyle3.ListRef.Levels;
            foreach (ListLevel listLev in Levels1)
            {
                listLev.UsePrevLevelPattern = true;
                listLev.NumberPrefix = "1.1.";
            }

            //Add Heading 3.
            for (int i = 0; i < 4; i++)
            {
                paragraph = section.AddParagraph();

                //Append text
                paragraph.AppendText(BuiltinStyle.Heading3.ToString());

                //Apply list style 3 for Heading 3
                paragraph.ApplyStyle(BuiltinStyle.Heading3);

                paragraph.ListFormat.ApplyStyle(listStyle3.Name);

            }

            // Specify the file name for the resulting Word document.
            String result = "Result-FormACatalogue.docx";

            // Save the Document object to a file in Docx format and dispose it.
            document.SaveToFile(result, FileFormat.Docx);
            document.Dispose();

            //Launch the MS Word file.
            WordDocViewer(result);
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
