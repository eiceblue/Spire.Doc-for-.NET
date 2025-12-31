using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

namespace InsertingImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a document
            Document document = new Document();

            //Add a seciton
            Section section = document.AddSection();

            //insert image
            InsertImage(section);

            //Save the file.
            document.SaveToFile("Sample.docx",FileFormat.Docx);

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");


        }

        private void InsertImage(Section section)
        {
            // Add a new paragraph to the section
			Paragraph paragraph = section.AddParagraph();
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

			// Load an image from file
			System.Drawing.Image ima = System.Drawing.Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png");

            // Append the image to the paragraph and set its width and height
            DocPicture picture = paragraph.AppendPicture(ima);
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
             DocPicture picture = paragraph.AppendPicture(@"..\..\..\..\..\..\Data\Spire.Doc.png");
            */

            picture.Width = 100;
			picture.Height = 100;

			// Add a new paragraph to the section
			paragraph = section.AddParagraph();
			paragraph.Format.LineSpacing = 20f;

			// Add text to the paragraph with specified formatting
			TextRange tr = paragraph.AppendText("Spire.Doc for .NET is a professional Word .NET library specially designed for developers to create, read, write, convert and print Word document files from any .NET (C#, VB.NET, ASP.NET) platform with fast and high-quality performance.");
			tr.CharacterFormat.FontName = "Arial";
			tr.CharacterFormat.FontSize = 14;

			// Add an empty paragraph to create spacing
			section.AddParagraph();

			// Add a new paragraph to the section
			paragraph = section.AddParagraph();
			paragraph.Format.LineSpacing = 20f;

			// Add text to the paragraph with specified formatting
			tr = paragraph.AppendText("As an independent Word .NET component, Spire.Doc for .NET doesn't need Microsoft Word to be installed on the machine. However, it can incorporate Microsoft Word document creation capabilities into any developers' .NET applications.");
			tr.CharacterFormat.FontName = "Arial";
			tr.CharacterFormat.FontSize = 14;
           
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
