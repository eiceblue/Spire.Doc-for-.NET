using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Spire.Doc;

namespace NestedMailMerage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a list to store DictionaryEntry objects
			List<DictionaryEntry> list = new List<DictionaryEntry>();

			// Create a DataSet object
			DataSet dsData = new DataSet();

			// Read XML data into the DataSet
			dsData.ReadXml(@"..\..\..\..\..\..\Data\Orders.xml");

			// Create a Document object
			Document document = new Document();

			// Load a Word document from file
			document.LoadFromFile(@"..\..\..\..\..\..\Data\NestedMailMerge.doc");

			// Create a DictionaryEntry for "Customer" with an empty value and add it to the list
			DictionaryEntry dictionaryEntry = new DictionaryEntry("Customer", string.Empty);
			list.Add(dictionaryEntry);

			// Create a DictionaryEntry for "Order" with a nested region condition and add it to the list
			dictionaryEntry = new DictionaryEntry("Order", "Customer_Id = %Customer.Customer_Id%");
			list.Add(dictionaryEntry);

			// Execute mail merge with nested regions using the DataSet and list of DictionaryEntry objects
			document.MailMerge.ExecuteWidthNestedRegion(dsData, list);

			// Save the merged document to a file 
			document.SaveToFile("Sample.docx", FileFormat.Docx);

			// Dispose the Document object 
			document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");
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
