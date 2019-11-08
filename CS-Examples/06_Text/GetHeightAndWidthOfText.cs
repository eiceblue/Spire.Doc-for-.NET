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

namespace GetHeightAndWidthOfText
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_2.docx");

            //Define the text string that we need to get the height and width.
            string text = "Your Office Development Master";

            //Finds and returns the string with formatting
            TextSelection selection = document.FindString(text, true, true);

            //Get the font
            Font font = selection.GetAsOneRange().CharacterFormat.Font;

            //Initialize graphics object
            Image fakeImage = new Bitmap(1, 1);
            Graphics graphics = Graphics.FromImage(fakeImage);

            //Measure string
            SizeF size = graphics.MeasureString(text, font);

            //Get the height and width of the text.
            StringBuilder content = new StringBuilder();
            content.AppendLine("text height: " + size.Height);
            content.AppendLine("text width: " + size.Width);

            String result = "Result-GetHeightAndWidthOfText.txt";

            //Save to file.
            File.WriteAllText(result,content.ToString());

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
