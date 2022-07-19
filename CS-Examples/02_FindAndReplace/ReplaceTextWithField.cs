using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ReplaceTextWithField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a new document
            Document document = new Document();
            //Load file from disk
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceTextWithField.docx");
            //Find the target text
            TextSelection selection= document.FindString("summary", false,true);
            //Get text range
            TextRange textRange=selection.GetAsOneRange();
            //Get it's owner paragraph
            Paragraph ownParagraph = textRange.OwnerParagraph;
            //Get the index of this text range
            int rangeIndex = ownParagraph.ChildObjects.IndexOf(textRange);
            //Remove the text range
            ownParagraph.ChildObjects.RemoveAt(rangeIndex);
            //Remove the objects which are behind the text range
            List<DocumentObject> tempList = new List<DocumentObject>();
            for(int i = rangeIndex; i < ownParagraph.ChildObjects.Count; i++)
            {
                //Add a copy of these objects into a temp list
                tempList.Add(ownParagraph.ChildObjects[rangeIndex].Clone());
                ownParagraph.ChildObjects.RemoveAt(rangeIndex);
            }
            //Append field to the paragraph
            ownParagraph.AppendField("MyFieldName", FieldType.FieldMergeField);
            //Put these objects back into the paragraph one by one
            foreach (DocumentObject obj in tempList)
            {
                ownParagraph.ChildObjects.Add(obj);
            }

            //Save the document
            string output = "ReplaceTextWithField_output.docx";
            document.SaveToFile(output,FileFormat.Docx);
            WordDocViewer(output);
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
