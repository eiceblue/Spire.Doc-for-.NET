using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ReplaceTextWithTable
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

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

            //Return TextSection by finding the key text string "Christmas Day, December 25".
            Section section = document.Sections[0];
            TextSelection selection = document.FindString("Christmas Day, December 25", true, true);

            //Return TextRange from TextSection, then get OwnerParagraph through TextRange.
            TextRange range = selection.GetAsOneRange();
            Paragraph paragraph = range.OwnerParagraph;

            //Return the zero-based index of the specified paragraph.
            Body body = paragraph.OwnerTextBody;
            int index = body.ChildObjects.IndexOf(paragraph);

            //Create a new table.
            Table table = section.AddTable(true);
            table.ResetCells(3, 3);

            //Remove the paragraph and insert table into the collection at the specified index.
            body.ChildObjects.Remove(paragraph);
            body.ChildObjects.Insert(index, table);

            String result = "Result-ReplaceTextWithTable.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

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
