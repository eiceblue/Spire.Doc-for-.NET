using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;
using Spire.Doc.Fields;
using System.Windows.Forms;

namespace ExtractTextFromTextBoxes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a Word document.
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractTextFromTextBoxes.docx");

            String result = "Result-ExtractTextFromTextBoxes.txt";

            //Verify whether the document contains a textbox or not.
            if (document.TextBoxes.Count > 0)
            {
                using (StreamWriter sw = File.CreateText(result))
                {
                    //Traverse the document.
                    foreach (Section section in document.Sections)
                    {
                        foreach (Paragraph p in section.Paragraphs)
                        {
                            foreach (DocumentObject obj in p.ChildObjects)
                            {
                                if (obj.DocumentObjectType == DocumentObjectType.TextBox)
                                {
                                    Spire.Doc.Fields.TextBox textbox = obj as Spire.Doc.Fields.TextBox;
                                    foreach (DocumentObject objt in textbox.ChildObjects)
                                    {
                                        //Extract text from paragraph in TextBox.
                                        if (objt.DocumentObjectType == DocumentObjectType.Paragraph)
                                        {
                                            sw.Write((objt as Paragraph).Text);
                                        }

                                        //Extract text from Table in TextBox.
                                        if (objt.DocumentObjectType == DocumentObjectType.Table)
                                        {
                                            Table table = objt as Table;
                                            ExtractTextFromTables(table, sw);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //Launch the result file.
            WordDocViewer(result);
        }

        //Extract text from Table .
        static void ExtractTextFromTables(Table table, StreamWriter sw)
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                TableRow row = table.Rows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    TableCell cell = row.Cells[j];
                    foreach (Paragraph paragraph in cell.Paragraphs)
                    {
                        sw.Write(paragraph.Text);
                    }
                }
            }
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
