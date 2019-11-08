using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace StartFromFormField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create the source document
            Document sourceDocument = new Document();

            //Load the source document from disk.
            sourceDocument.LoadFromFile(@"..\..\..\..\..\..\Data\TextInputField.docx");

            //Create a destination document
            Document destinationDoc = new Document();

            //Add a section
            Section section = destinationDoc.AddSection();

            //Define a variables
            int index = 0;

            //Traverse FormFields
            foreach (FormField field in sourceDocument.Sections[0].Body.FormFields)
            {
                //Find FieldFormTextInput type field
                if (field.Type == FieldType.FieldFormTextInput)
                {
                    //Get the paragraph
                    Paragraph paragraph = field.OwnerParagraph;

                    //Get the index
                    index = sourceDocument.Sections[0].Body.ChildObjects.IndexOf(paragraph);
                    break;
                }
            }

            //Extract the content
            for (int i = index; i < index + 3; i++)
            {
                //Clone the ChildObjects of source document
                DocumentObject doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone();

                //Add to destination document 
                section.Body.ChildObjects.Add(doobj);
            }

            //Save the document.
            destinationDoc.SaveToFile("FromFormField.docx", FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer("FromFormField.docx");
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
