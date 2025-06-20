using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Interface;
using Spire.Doc.Reporting;

namespace MergeEventHandler
{
    public partial class Form1 : Form
    {
        private int lastIndex = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document instance
			Document document = new Document();

			// Load the document from the specified file
			document.LoadFromFile(@"..\..\..\..\..\..\Data\MergeEventHandler.doc");

			// Create a list of CustomerRecord objects
			List<CustomerRecord> customerRecords = new List<CustomerRecord>();

			// Add customer records to the list
			CustomerRecord c1 = new CustomerRecord();
			c1.ContactName = "Lucy";
			c1.Fax = "786-324-10";
			c1.Date = DateTime.Now;
			customerRecords.Add(c1);

			CustomerRecord c2 = new CustomerRecord();
			c2.ContactName = "Lily";
			c2.Fax = "779-138-13";
			c2.Date = DateTime.Now;
			customerRecords.Add(c2);

			CustomerRecord c3 = new CustomerRecord();
			c3.ContactName = "James";
			c3.Fax = "363-287-02";
			c3.Date = DateTime.Now;
			customerRecords.Add(c3);

			// Subscribe to the MergeField event
			document.MailMerge.MergeField += new MergeFieldEventHandler(MailMerge_MergeField);

			// Execute the mail merge using the customerRecords list as the data source
			document.MailMerge.ExecuteGroup(new MailMergeDataTable("Customer", customerRecords));

			// Save the merged document to a file
			document.SaveToFile(@"Sample.doc", FileFormat.Doc);

			// Dispose the document
			document.Dispose();

            //Launching the MS Word file.
            WordDocViewer(@"Sample.doc");

        }

        void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
        {
            // Check if the current row index is greater than the lastIndex
			if (args.RowIndex > lastIndex)
			{
				// Update the lastIndex with the current row index
				lastIndex = args.RowIndex;

				// Add a page break before the current merge field
				AddPageBreakForMergeField(args.CurrentMergeField);
			}
        }

        void AddPageBreakForMergeField(IMergeField mergeField)
        {
            //Find position of needing to add page break
            bool foundGroupStart = false;
			Paragraph paragraph = mergeField.PreviousSibling.Owner as Paragraph;
			MergeField previousMergeField = null;

			// Find the group start merge field by traversing the previous sibling paragraphs
			while (!foundGroupStart)
			{
				paragraph = paragraph.PreviousSibling as Paragraph;

				for (int i = 0; i < paragraph.Items.Count; i++)
				{
					previousMergeField = paragraph.Items[i] as MergeField;

					if ((previousMergeField != null) && (previousMergeField.Prefix == "GroupStart"))
					{
						foundGroupStart = true;
						break;
					}
				}
			}

			// Append a page break to the paragraph
			paragraph.AppendBreak(BreakType.PageBreak);
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

    public class CustomerRecord
    {
        private string m_contactName;
        public string ContactName
        {
            get
            {
                return m_contactName;
            }
            set
            {
                m_contactName = value;
            }
        }

        private string m_fax;
        public string Fax
        {
            get
            {
                return m_fax;
            }
            set
            {
                m_fax = value;
            }
        }

        private DateTime m_date;
        public DateTime Date
        {
            get
            {
                return m_date;
            }
            set
            {
                m_date = value;
            }
        }
    }
}
