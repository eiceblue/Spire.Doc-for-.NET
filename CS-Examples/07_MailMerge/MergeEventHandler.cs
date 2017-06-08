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
            //Create word document
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Fax2.doc");
            lastIndex = 0;

            List<CustomerRecord> customerRecords = new List<CustomerRecord>();
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

            //Execute mailmerge
            document.MailMerge.MergeField += new MergeFieldEventHandler(MailMerge_MergeField);
            document.MailMerge.ExecuteGroup(new MailMergeDataTable("Customer", customerRecords));

            //Save doc file.
            document.SaveToFile(@"Sample.doc", FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer(@"Sample.doc");

        }

        void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
        {
            //Next row
            if (args.RowIndex > lastIndex)
            {
                lastIndex = args.RowIndex;
                AddPageBreakForMergeField(args.CurrentMergeField);
            }
        }

        void AddPageBreakForMergeField(IMergeField mergeField)
        {
            //Find position of needing to add page break
            bool foundGroupStart = false;
            Paragraph paramgraph = mergeField.PreviousSibling.Owner as Paragraph;
            MergeField merageField = null;
            while (!foundGroupStart)
            {
                paramgraph = paramgraph.PreviousSibling as Paragraph;
                for (int i = 0; i < paramgraph.Items.Count; i++)
                {
                    merageField = paramgraph.Items[i] as MergeField;
                    if ((merageField != null) && (merageField.Prefix == "GroupStart"))
                    {
                        foundGroupStart = true;
                        break;
                    }
                }
            }

            paramgraph.AppendBreak(BreakType.PageBreak);
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
