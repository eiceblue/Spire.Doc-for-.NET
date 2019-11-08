using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

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
            //Create Word document.
            Document document = new Document();

            //Add a new section. 
            Section section = document.AddSection();
            Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

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
            ListStyle listSty2 = new ListStyle(document, ListType.Numbered);
            foreach (ListLevel listLev in listSty2.Levels)
            {
                listLev.UsePrevLevelPattern = true;
                listLev.NumberPrefix = "1.";
            }
            listSty2.Name = "MyStyle2";
            document.ListStyles.Add(listSty2);
            paragraph.ListFormat.ApplyStyle(listSty2.Name);

            //Add list style 3.
            ListStyle listSty3 = new ListStyle(document, ListType.Numbered);
            foreach (ListLevel listLev in listSty3.Levels)
            {
                listLev.UsePrevLevelPattern = true;
                listLev.NumberPrefix = "1.1.";
            }
            listSty3.Name = "MyStyle3";
            document.ListStyles.Add(listSty3);

            //Add Heading 3.
            for (int i = 0; i < 4; i++)
            {
                paragraph = section.AddParagraph();

                //Append text
                paragraph.AppendText(BuiltinStyle.Heading3.ToString());

                //Apply list style 3 for Heading 3
                paragraph.ApplyStyle(BuiltinStyle.Heading3);
                paragraph.ListFormat.ApplyStyle(listSty3.Name);
            }

            String result = "Result-FormACatalogue.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx);

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
