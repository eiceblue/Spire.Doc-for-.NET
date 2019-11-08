using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ConvertFieldToBodyText
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

            //Traverse FormFields
            foreach (FormField field in sourceDocument.Sections[0].Body.FormFields)
            {
                //Find FieldFormTextInput type field
                if (field.Type == FieldType.FieldFormTextInput)
                {
                    //Get the paragraph
                    Paragraph paragraph = field.OwnerParagraph;

                    //Define variables
                    int startIndex = 0;
                    int endIndex = 0;

                    //Create a new TextRange
                    TextRange textRange = new TextRange(sourceDocument);

                    //Set text for textRange
                    textRange.Text = paragraph.Text;

                    //Traverse DocumentObjectS of field paragraph
                    foreach (DocumentObject obj in paragraph.ChildObjects)
                    {
                        //If its DocumentObjectType is BookmarkStart
                        if (obj.DocumentObjectType == DocumentObjectType.BookmarkStart)
                        {
                            //Get the index
                            startIndex = paragraph.ChildObjects.IndexOf(obj);
                        }
                        //If its DocumentObjectType is BookmarkEnd
                        if (obj.DocumentObjectType == DocumentObjectType.BookmarkEnd)
                        {
                            //Get the index
                            endIndex = paragraph.ChildObjects.IndexOf(obj);
                        }
                    }
                    //Remove ChildObjects
                    for (int i = endIndex; i > startIndex; i--)
                    {
                        //If it is TextFormField
                        if (paragraph.ChildObjects[i] is TextFormField)
                        {
                            TextFormField textFormField = paragraph.ChildObjects[i] as TextFormField;

                            //Remove the field object
                            paragraph.ChildObjects.Remove(textFormField);
                        }
                        else
                        {
                            paragraph.ChildObjects.RemoveAt(i);
                        }
                    }
                    //Insert the new TextRange
                    paragraph.ChildObjects.Insert(startIndex, textRange);
                    break;
                }

            }
            //Save the document.
            sourceDocument.SaveToFile("Output.docx", FileFormat.Docx);

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
