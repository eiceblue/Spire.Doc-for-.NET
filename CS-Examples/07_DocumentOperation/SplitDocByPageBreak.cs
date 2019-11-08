using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SplitDocByPageBreak
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
            Document original = new Document();

            //Load the file from disk.
            original.LoadFromFile(@"..\..\..\..\..\..\..\Data\SplitWordFileByPageBreak.docx");

            //Create a new word document and add a section to it.
            Document newWord = new Document();
            Section section = newWord.AddSection();
            original.CloneDefaultStyleTo(newWord);
            original.CloneThemesTo(newWord);
            original.CloneCompatibilityTo(newWord);
          
            //Split the original word document into separate documents according to page break.
            int index = 0;

            //Traverse through all sections of original document.
            foreach (Section sec in original.Sections)
            {
                //Traverse through all body child objects of each section.
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    if (obj is Paragraph)
                    {
                        Paragraph para = obj as Paragraph;
                        sec.CloneSectionPropertiesTo(section); 
                        //Add paragraph object in original section into section of new document.
                        section.Body.ChildObjects.Add(para.Clone());

                        foreach (DocumentObject parobj in para.ChildObjects)
                        {
                            if (parobj is Break && (parobj as Break).BreakType == BreakType.PageBreak)
                            {
                                //Get the index of page break in paragraph.
                                int i = para.ChildObjects.IndexOf(parobj);

                                //Remove the page break from its paragraph.
                                section.Body.LastParagraph.ChildObjects.RemoveAt(i);

                                //Save the new document to a Docx file.
                                newWord.SaveToFile(String.Format("Result-SplitWordFileByPageBreak-{0}.docx", index), FileFormat.Docx);
                                index++;

                                //Create a new document and add a section.
                                newWord = new Document();
                                section = newWord.AddSection();
                                original.CloneDefaultStyleTo(newWord);
                                original.CloneThemesTo(newWord);
                                original.CloneCompatibilityTo(newWord);
                                sec.CloneSectionPropertiesTo(section);
                                //Add paragraph object in original section into section of new document.
                                section.Body.ChildObjects.Add(para.Clone());
                                if (section.Paragraphs[0].ChildObjects.Count == 0)
                                {
                                    //Remove the first blank paragraph.
                                    section.Body.ChildObjects.RemoveAt(0);
                                }
                                else
                                {
                                    //Remove the child objects before the page break.
                                    while (i >= 0)
                                    {
                                        section.Paragraphs[0].ChildObjects.RemoveAt(i);
                                        i--;
                                    }
                                }
                            }
                        }
                    }
                    if (obj is Table)
                    {
                        //Add table object in original section into section of new document.
                        section.Body.ChildObjects.Add(obj.Clone());
                    }
                }
            }

            //Save to file.
            String result = String.Format("Result-SplitWordFileByPageBreak-{0}.docx", index);
            newWord.SaveToFile(result, FileFormat.Docx2013);

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
