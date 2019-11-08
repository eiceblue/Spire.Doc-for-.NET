using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Text.RegularExpressions;

namespace ReplaceContentWithDoc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create the first document
            Document document1 = new Document();

            //Load the first document from disk.
            document1.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceContentWithDoc.docx");

            //Create the second document
            Document document2 = new Document();

            //Load the second document from disk.
            document2.LoadFromFile(@"..\..\..\..\..\..\Data\Insert.docx");

            //Get the first section of the first document 
            Section section1 = document1.Sections[0];

            //Create a regex
            Regex regex = new Regex(@"\[MY_DOCUMENT\]", RegexOptions.None);

            //Find the text by regex
            TextSelection[] textSections = document1.FindAllPattern(regex);

            //Travel the found strings
            foreach (TextSelection seletion in textSections)
            {

                //Get the para
                Paragraph para = seletion.GetAsOneRange().OwnerParagraph;

                //Get textRange
                TextRange textRange = seletion.GetAsOneRange();

                //Get the para index
                int index = section1.Body.ChildObjects.IndexOf(para);

                //Insert the paragraphs of document2
                foreach (Section section2 in document2.Sections)
                {
                    foreach (Paragraph paragraph in section2.Paragraphs)
                    {
                        section1.Body.ChildObjects.Insert(index, paragraph.Clone() as Paragraph);
                    }
                }
                //Remove the found textRange
                para.ChildObjects.Remove(textRange);
            }

            //Save the document.
            document1.SaveToFile("Output.docx", FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer("Output.docx");
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
