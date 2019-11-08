using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Fields;
using Spire.Doc.Documents;
namespace HandleAskField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //Create and load Word document.
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\HandleAskField.docx");

            //call UpdateFieldsHandler event to handle the ASK field.
            doc.UpdateFields += new UpdateFieldsHandler(doc_UpdateFields);
            //update the fields in the document.
            doc.IsUpdateFields = true;        
            //save the document.
            doc.SaveToFile("HandleAskField.docx", FileFormat.Docx);
            WordDocViewer("HandleAskField.docx");
         
        }
        private static void doc_UpdateFields(object sender, IFieldsEventArgs args)
        {     
            if (args is AskFieldEventArgs)
            {
                AskFieldEventArgs askArgs = args as AskFieldEventArgs;
                
                if (askArgs.BookmarkName == "1")
                {
                    askArgs.ResponseText = "Thank you. This is my first time to come to a Chinese restaurant. Could you tell me the different features of Chinese food?";
                }
                
                if (askArgs.BookmarkName == "2")
                {
                    askArgs.ResponseText = "No more, thank you. I'm quite full.";
                }
            }
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
