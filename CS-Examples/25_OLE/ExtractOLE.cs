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
using System.IO;

namespace ExtractOLE
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create document and load file from disk
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\OLEs.docx");

            //Traverse through all sections of the word document    
            foreach (Section sec in doc.Sections)
            {
                //Traverse through all Child Objects in the body of each section
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    //find the paragraph
                    if (obj is Paragraph)
                    {
                        Paragraph par = obj as Paragraph;
                        foreach (DocumentObject o in par.ChildObjects)
                        {
                            //check whether the object is OLE
                            if (o.DocumentObjectType == DocumentObjectType.OleObject)
                            {
                                DocOleObject Ole = o as DocOleObject;
                                string s = Ole.ObjectType;

                                //check whether the object type is "Acrobat.Document.11"
                                if (s == "AcroExch.Document.DC")
                                {
                                    //write the data of OLE into file
                                    File.WriteAllBytes("Result.pdf", Ole.NativeData);
                                    FileViewer("Result.pdf");
                                }

                                //check whether the object type is "Excel.Sheet.8"
                                else if (s == "Excel.Sheet.8")
                                {
                                    File.WriteAllBytes("ExcelResult.xls", Ole.NativeData);
                                    FileViewer("ExcelResult.xls");
                                }

                                //check whether the object type is "PowerPoint.Show.12"
                                else if (s == "PowerPoint.Show.12")
                                {
                                    File.WriteAllBytes("PPTResult.pptx", Ole.NativeData);
                                    FileViewer("PPTResult.pptx");
                                }
                            }
                        }
                    }
                }
            }
        }
        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
