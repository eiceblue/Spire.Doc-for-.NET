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
            // Create a new Document object.
            Document document = new Document();

            // Add a Section to the document.
            Section section = document.AddSection();

            // Get the first Paragraph of the Section, or add a new Paragraph if none exists.
            Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

            // Add a new Paragraph to the Section.
            paragraph = section.AddParagraph();

            // Set the text content of the Paragraph and apply the Heading1 style.
            paragraph.AppendText(BuiltinStyle.Heading1.ToString());
            paragraph.ApplyStyle(BuiltinStyle.Heading1);

            // Apply a numbered list format to the Paragraph.
            paragraph.ListFormat.ApplyNumberedStyle();

            // Add another Paragraph to the Section.
            paragraph = section.AddParagraph();

            // Set the text content of the Paragraph and apply the Heading2 style.
            paragraph.AppendText(BuiltinStyle.Heading2.ToString());
            paragraph.ApplyStyle(BuiltinStyle.Heading2);

            // Create a new ListStyle object with the Numbered list type.
            ListStyle listSty2 = new ListStyle(document, ListType.Numbered);

            // Iterate over the levels of the ListStyle and customize them.
            foreach (ListLevel listLev in listSty2.Levels)
            {
                listLev.UsePrevLevelPattern = true;
                listLev.NumberPrefix = "1.";
            }

            // Set the name of the ListStyle and add it to the document's ListStyles collection.
            listSty2.Name = "MyStyle2";
            document.ListStyles.Add(listSty2);

            // Apply the ListStyle to the current Paragraph.
            paragraph.ListFormat.ApplyStyle(listSty2.Name);

            // Create another ListStyle object with the Numbered list type.
            ListStyle listSty3 = new ListStyle(document, ListType.Numbered);

            // Iterate over the levels of the ListStyle and customize them.
            foreach (ListLevel listLev in listSty3.Levels)
            {
                listLev.UsePrevLevelPattern = true;
                listLev.NumberPrefix = "1.1.";
            }

            // Set the name of the ListStyle and add it to the document's ListStyles collection.
            listSty3.Name = "MyStyle3";
            document.ListStyles.Add(listSty3);

            // Add four Paragraphs to the Section and apply Heading3 style and ListStyle to each.
            for (int i = 0; i < 4; i++)
            {
                paragraph = section.AddParagraph();
                paragraph.AppendText(BuiltinStyle.Heading3.ToString());
                paragraph.ApplyStyle(BuiltinStyle.Heading3);
                paragraph.ListFormat.ApplyStyle(listSty3.Name);
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
