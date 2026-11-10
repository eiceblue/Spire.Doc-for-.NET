using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Documents;

namespace CreatePictureBullet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {  
            // Create a new Document object to represent a Word document.
            Document document = new Document();

            // Add a new section to the document. A section is a fundamental structural unit in Word.
            Section sec = document.AddSection();

            // Add a paragraph to the section and assign it to the variable 'paragraph'.
            Spire.Doc.Documents.Paragraph paragraph = sec.AddParagraph();

            // Create a bulleted list style named "bulletList".
            ListStyle listStyle = document.Styles.Add(ListType.Bulleted, "bulletList");

            // Get the collection of list levels (up to 9 levels, indexed from 0) for this list style.
            ListLevelCollection Levels = listStyle.ListRef.Levels;

            // Enable picture bullet for level 0 (the top-level list items).
            Levels[0].CreatePictureBullet();

            // Load an image file as the bullet for level 0 using a relative path.
            Levels[0].PictureBullet.LoadImage(@"..\..\..\..\..\..\Data\Word.png");

            // Enable picture bullet for level 1 (second-level indented list items).
            Levels[1].CreatePictureBullet();

            // Load another image file as the bullet for level 1 using a relative path.
            Levels[1].PictureBullet.LoadImage(@"..\..\..\..\..\..\Data\logo.png");

            // Add a new paragraph for the first-level list item.
            paragraph = sec.AddParagraph();

            // Append the text "List Item 1" to the paragraph.
            paragraph.AppendText("List Item 1");

            // Apply the custom "bulletList" style to this paragraph.
            paragraph.ListFormat.ApplyStyle(listStyle);

            // Add another paragraph for the second-level list item.
            paragraph = sec.AddParagraph();

            // Append the text "List Item 1.1" to this paragraph.
            paragraph.AppendText("List Item 1.1");

            // Apply the same list style to make it part of the list structure.
            paragraph.ListFormat.ApplyStyle(listStyle);

            // Set the list level of this paragraph to 1 (i.e., second indentation level).
            paragraph.ListFormat.ListLevelNumber = 1;

            // Save the document as "CreatePictureBullet.docx" in Word 2016 (.docx) format.
            document.SaveToFile("CreatePictureBullet.docx", FileFormat.Docx2016);

            // Close the document to release internal resources such as file handles.
            document.Close();

            // Explicitly dispose of the document object to ensure memory is freed immediately.
            document.Dispose();


            //Launching the Word file.
            WordDocViewer("CreatePictureBullet.docx");
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
