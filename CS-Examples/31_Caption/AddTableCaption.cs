using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddTableCaption
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create word document
            Document document = new Document();
            //Load file
            document.LoadFromFile(@"..\..\..\..\..\..\Data\TableTemplate.docx");

            //Get the first table
            Body body = document.Sections[0].Body;
            Table table = body.Tables[0] as Table;

            //Add caption to the table
            table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem);

            //Update fields
            document.IsUpdateFields = true;

            //Save the file
            string output = "AddTableCaption_result.docx";
            document.SaveToFile(output,FileFormat.Docx);

            //Launching the file
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
