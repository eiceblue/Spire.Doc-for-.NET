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
        
			// Create a new document
			Document doc = new Document();

			// Load the document from a file
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\HandleAskField.docx");

			// Subscribe to the UpdateFields event
			doc.UpdateFields += new UpdateFieldsHandler(doc_UpdateFields);

			// Enable field update
			doc.IsUpdateFields = true;

			// Save the modified document to a file
			doc.SaveToFile("HandleAskField.docx", FileFormat.Docx);

			// Dispose the document object
			doc.Dispose();
			
            WordDocViewer("HandleAskField.docx");
         
        }
		// Event handler for updating fields
		private static void doc_UpdateFields(object sender, IFieldsEventArgs args)
		{     
			// Check if the event arguments are of type AskFieldEventArgs
			if (args is AskFieldEventArgs)
			{
				AskFieldEventArgs askArgs = args as AskFieldEventArgs;
				
				// Handle different bookmarks and set response text accordingly
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
