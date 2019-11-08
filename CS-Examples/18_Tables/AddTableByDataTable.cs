using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AddTableByDataTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a Word document
            Document document = new Document();

            //Get the first section
            Section section = document.AddSection();

            //Add paragraph style
            ParagraphStyle style = new ParagraphStyle(document);
            style.CharacterFormat.FontSize = 20f;
            style.CharacterFormat.Bold = true;
            style.CharacterFormat.TextColor = Color.CadetBlue;
            document.Styles.Add(style);

            //Create a paragraph and append text
            Paragraph para = section.AddParagraph();
            para.AppendText("Table");
            //Apply style
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            para.ApplyStyle(style.Name);

            //Load data
            DataSet ds = new DataSet();
            ds.ReadXml(@"..\..\..\..\..\..\Data\dataTable.xml");

            //Get the first data table
            DataTable dataTable = ds.Tables[0];

            //Add a table
            Table table = section.AddTable(true);
            //Set its width
            table.PreferredWidth = new PreferredWidth(WidthType.Percentage, 100);

            //Fill table with the data of datatable
            FillTableUsingDataTable(table, dataTable);

            //Set table style
            table.TableFormat.Paddings.All = 5;
            table.FirstRow.RowFormat.BackColor = Color.CadetBlue;

            //Save the Word file
            string output = "AddTableUsingDataTable_out.docx";
            document.SaveToFile(output, FileFormat.Docx2013);

            //Launch the file
            FileViewer(output);
        }
        private static void FillTableUsingDataTable(Table table, DataTable dataTable)
        {
            int columnCount = dataTable.Columns.Count;

            foreach (DataRow dataRow in dataTable.Rows)
            {
                TableRow row = table.AddRow(columnCount);
                foreach (DataColumn dataColumn in dataTable.Columns)
                {
                    int columnIndex = dataTable.Columns.IndexOf(dataColumn);
                    string value = dataRow[dataColumn].ToString();
                    TableCell cell = row.Cells[columnIndex];
                    //Add paragraph for cell
                    Paragraph para = cell.AddParagraph();
                    //Append text from datatable
                    para.AppendText(value);
                    //Set the alignment of cell
                    cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                }
            }
        }
        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
