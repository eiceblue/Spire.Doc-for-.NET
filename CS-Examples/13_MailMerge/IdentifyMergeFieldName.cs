using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;

namespace IdentifyMergeFieldName
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\IdentifyMergeFieldNames.docx");

            //Get the collection of group names.
            string[] GroupNames = document.MailMerge.GetMergeGroupNames();

            //Get the collection of merge field names in a specific group.
            string[] MergeFieldNamesWithinRegion = document.MailMerge.GetMergeFieldNames("Products");

            //Get the collection of all the merge field names.
            string[] MergeFieldNames = document.MailMerge.GetMergeFieldNames();

            StringBuilder content = new StringBuilder();
            content.AppendLine("----------------Group Names-----------------------------------------");
            for (int i = 0; i < GroupNames.Length; i++)
            {
                content.AppendLine(GroupNames[i]);
            }

            content.AppendLine("----------------Merge field names within a specific group-----------");
            for (int j = 0; j < MergeFieldNamesWithinRegion.Length; j++)
            {
                content.AppendLine(MergeFieldNamesWithinRegion[j]);
            }

            content.AppendLine("----------------All of the merge field names------------------------");
            for (int k = 0; k < MergeFieldNames.Length; k++)
            {
                content.AppendLine(MergeFieldNames[k]);
            }

            String result = "Result-IdentifyMergeFieldNames.txt";

            //Save to file.
            File.WriteAllText(result,content.ToString());
			
			// Dispose the document object
			document.Dispose();
			
            //Launch the file.
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
