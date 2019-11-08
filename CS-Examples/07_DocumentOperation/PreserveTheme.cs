using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace PreserveTheme
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load the source document
            string input = @"..\..\..\..\..\..\..\Data\Theme.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Create a new Word document
            Document newWord = new Document();
            //Clone default style, theme, compatibility from the source file to the destination document
            doc.CloneDefaultStyleTo(newWord);
            doc.CloneThemesTo(newWord);
            doc.CloneCompatibilityTo(newWord);

            //Add the cloned section to destination document
            newWord.Sections.Add(doc.Sections[0].Clone());

            //Save and launch document
            string output = "PreserveTheme.docx";
            newWord.SaveToFile(output, FileFormat.Docx);
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
