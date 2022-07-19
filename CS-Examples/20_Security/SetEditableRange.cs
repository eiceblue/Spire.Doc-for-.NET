using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetEditableRange
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
            //Load file from disk
            document.LoadFromFile(@"..\..\..\..\..\..\Data\SetEditableRange.docx");
            //Protect whole document
            document.Protect(ProtectionType.AllowOnlyReading,"password");
            //Create tags for permission start and end
            PermissionStart start = new PermissionStart(document,"testID");
            PermissionEnd end = new PermissionEnd(document, "testID");
            //Add the start and end tags to allow the first paragraph to be edited.
            document.Sections[0].Paragraphs[0].ChildObjects.Insert(0, start);
            document.Sections[0].Paragraphs[0].ChildObjects.Add(end);
            //Save the document
            string output = "SetEditableRange_output.docx";
            document.SaveToFile(output,FileFormat.Docx);
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
