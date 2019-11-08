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

namespace CountWordsNumber
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

            //Count the number of words.
            StringBuilder content = new StringBuilder();
            content.AppendLine("CharCount: " + document.BuiltinDocumentProperties.CharCount);
            content.AppendLine("CharCountWithSpace: " + document.BuiltinDocumentProperties.CharCountWithSpace);
            content.AppendLine("WordCount: " + document.BuiltinDocumentProperties.WordCount);

            //Save to file.
            String result = "Result-CountWordsNumber.txt";
            File.WriteAllText(result, content.ToString());

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
