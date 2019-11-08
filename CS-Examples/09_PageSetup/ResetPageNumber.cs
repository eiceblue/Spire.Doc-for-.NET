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

namespace ResetPageNumber
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create three Word documents and load three different Word documents from disk.
            Document document1 = new Document();
            document1.LoadFromFile(@"..\..\..\..\..\..\Data\ResetPageNumber1.docx");

            Document document2 = new Document();
            document2.LoadFromFile(@"..\..\..\..\..\..\Data\ResetPageNumber2.docx");

            Document document3 = new Document();
            document3.LoadFromFile(@"..\..\..\..\..\..\Data\ResetPageNumber3.docx");

            //Use section method to combine all documents into one word document.
            foreach (Section sec in document2.Sections)
            {
                document1.Sections.Add(sec.Clone());
            }
            foreach (Section sec in document3.Sections)
            {
                document1.Sections.Add(sec.Clone());
            }

            //Traverse every section of document1.
            foreach (Section sec in document1.Sections)
            {
                //Traverse every object of the footer.
                foreach (DocumentObject obj in sec.HeadersFooters.Footer.ChildObjects)
                {
                    if (obj.DocumentObjectType == DocumentObjectType.StructureDocumentTag)
                    {
                        DocumentObject para = obj.ChildObjects[0];
                        foreach (DocumentObject item in para.ChildObjects)
                        {
                            if (item.DocumentObjectType == DocumentObjectType.Field)
                                //Find the item and its field type is FieldNumPages.
                                if ((item as Field).Type == FieldType.FieldNumPages)
                                {
                                    //Change field type to FieldSectionPages.
                                    (item as Field).Type = FieldType.FieldSectionPages;
                                }
                        }
                    }
                }
            }

            //Restart page number of section and set the starting page number to 1.
            document1.Sections[1].PageSetup.RestartPageNumbering = true;
            document1.Sections[1].PageSetup.PageStartingNumber = 1;

            document1.Sections[2].PageSetup.RestartPageNumbering = true;
            document1.Sections[2].PageSetup.PageStartingNumber = 1;

            String result = "Result-ResetPageNumber.docx";

            //Save to file.
            document1.SaveToFile(result, FileFormat.Docx2013);

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
