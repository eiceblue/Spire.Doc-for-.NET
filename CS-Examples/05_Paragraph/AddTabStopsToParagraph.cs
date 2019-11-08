using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddTabStopsToParagraph
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

            //Add a section.
            Section section = document.AddSection();

            //Add paragraph 1.
            Paragraph paragraph1 = section.AddParagraph();

            //Add tab and set its position (in points).
            Tab tab = paragraph1.Format.Tabs.AddTab(28);

            //Set tab alignment.
            tab.Justification = TabJustification.Left;

            //Move to next tab and append text.
            paragraph1.AppendText("\tWashing Machine");

            //Add another tab and set its position (in points).
            tab = paragraph1.Format.Tabs.AddTab(280);

            //Set tab alignment.
            tab.Justification = TabJustification.Left;

            //Specify tab leader type.
            tab.TabLeader = TabLeader.Dotted;

            //Move to next tab and append text.
            paragraph1.AppendText("\t$650");

            //Add paragraph 2.
            Paragraph paragraph2 = section.AddParagraph();

            //Add tab and set its position (in points).
            tab = paragraph2.Format.Tabs.AddTab(28);

            //Set tab alignment.
            tab.Justification = TabJustification.Left;

            //Move to next tab and append text.
            paragraph2.AppendText("\tRefrigerator"); 

            //Add another tab and set its position (in points).
            tab = paragraph2.Format.Tabs.AddTab(280);

            //Set tab alignment.
            tab.Justification = TabJustification.Left;

            //Specify tab leader type.
            tab.TabLeader = TabLeader.NoLeader;

            //Move to next tab and append text.
            paragraph2.AppendText("\t$800");

            String result = "Result-AddTabStopsToParagraph.docx";

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
