using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;
using Spire.Doc.Fields;
using System.Data;
using System.Data.OleDb;
using System.Xml.Linq;
using System.Linq;
using Spire.Doc.Reporting;
using System.Collections;
using System.Collections.Generic;
namespace AlternateRows
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String inputDataBase = @"..\..\..\..\..\..\Data\demo.mdb";
            String input = @"..\..\..\..\..\..\Data\ExecuteWithDataTable.doc";

            // Get a dataTable
            DataTable orderTable = GetCountryDataTable(inputDataBase);
            Document doc = new Document();
            //Load a mail merge template file
            doc.LoadFromFile(input);
          
            doc.MailMerge.MergeField += new MergeFieldEventHandler(MailMerge_MergeField);
            //Fill mergedField with data from dataTable
            doc.MailMerge.ExecuteWidthRegion(orderTable);

            string result = "AlternateRows_out.doc";
            doc.SaveToFile(result, FileFormat.Doc);
            WordViewer(result);
        }
        int rowIndex = 0;
        void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
        {
            // Catch the beginning of a new row.
            if (args.CurrentMergeField.FieldName.Equals("Name"))
            {
                // Set the color depending on whether the row number is even or odd.
                Color rowColor;
                if (rowIndex % 2 == 0)
                    rowColor = Color.FromArgb(215, 227, 235);
                else
                    rowColor = Color.FromArgb(240, 242, 242);

                TableCell cell = args.CurrentMergeField.OwnerParagraph.Owner as TableCell;
                TableRow row = cell.OwnerRow;

                row.RowFormat.BackColor = rowColor;
                rowIndex++;
            }
        }
    
        private DataTable GetCountryDataTable(string inputDataBase)
        {
            // Open a database connection
            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + inputDataBase;
            OleDbConnection connection = new OleDbConnection(connString);
            connection.Open();

            // Create the SQL command.
            string commandString = "SELECT * FROM Country";
            OleDbCommand command = new OleDbCommand(commandString, connection);

            // Create the data adapter.
            OleDbDataAdapter adapter = new OleDbDataAdapter(command);

            // Fill the results from the database into a DataTable.
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);
            dataTable.TableName = "Country";
            connection.Close();

            return dataTable;
        }
        private void WordViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
