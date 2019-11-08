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

namespace InsertPageBreakFirstApproach
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_2.docx");

            //Find the specified word "technology" where we want to insert the page break.
            TextSelection[] selections = document.FindAllString("technology", true, true);

            //Traverse each word "technology".
            foreach (TextSelection ts in selections)
            {
                TextRange range = ts.GetAsOneRange();
                Paragraph paragraph = range.OwnerParagraph;
                int index = paragraph.ChildObjects.IndexOf(range);

                //Create a new instance of page break and insert a page break after the word "technology".
                Break pageBreak = new Break(document, BreakType.PageBreak);
                paragraph.ChildObjects.Insert(index + 1, pageBreak);
            }

            String result = "Result-InsertPageBreakAtSpecifiedLocation.docx";

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
