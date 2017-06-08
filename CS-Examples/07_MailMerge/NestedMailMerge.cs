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
            List<DictionaryEntry> list = new List<DictionaryEntry>();
            DataSet dsData = new DataSet();

            dsData.ReadXml(@"..\..\..\..\..\..\Data\Orders.xml");

            //Create word document
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Invoice.doc");

            DictionaryEntry dictionaryEntry = new DictionaryEntry("Customer", string.Empty);
            list.Add(dictionaryEntry);

            dictionaryEntry = new DictionaryEntry("Order", "Customer_Id = %Customer.Customer_Id%");
            list.Add(dictionaryEntry);

            document.MailMerge.ExecuteWidthNestedRegion(dsData, list);
          
            //Save doc file.
            document.SaveToFile("Sample.doc", FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer("Sample.doc");
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
