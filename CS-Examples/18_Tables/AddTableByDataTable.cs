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

			//Create a new ParagraphStyle instance
			ParagraphStyle style = new ParagraphStyle(document);

			//Set the CharacterFormat of the style
			style.CharacterFormat.FontSize = 20f;
			style.CharacterFormat.Bold = true;
			style.CharacterFormat.TextColor = Color.CadetBlue;

			//Add the style to document
			document.Styles.Add(style);

			//Create a paragraph 
			Paragraph para = section.AddParagraph();

			//Append text
			para.AppendText("Table");

			//Set horizontal alignment for the paragraph
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			//Apply the new style
			para.ApplyStyle(style.Name);

			//Create a DataSet instance
			DataSet ds = new DataSet();

			//Load data from a xml file
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
			table.Format.Paddings.All = 5;

			for (int i = 0; i < table.FirstRow.Cells.Count; i++)
			{
				table.FirstRow.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.CadetBlue;
			}

            //Save the Word file
            string output = "AddTableUsingDataTable_out.docx";
			document.SaveToFile(output, FileFormat.Docx2013);
			
			//Dispose the document
			document.Dispose();

            //Launch the file
            FileViewer(output);
        }
        private static void FillTableUsingDataTable(Table table, DataTable dataTable)
        {
            //Get the count of the columns
			int columnCount = dataTable.Columns.Count;

			//Loop through the rows of data table
			foreach (DataRow dataRow in dataTable.Rows)
			{
				TableRow row = table.AddRow(columnCount);
				foreach (DataColumn dataColumn in dataTable.Columns)
				{

					//Get the column index
					int columnIndex = dataTable.Columns.IndexOf(dataColumn);

					//Get the value 
					string value = dataRow[dataColumn].ToString();

					//Get the cell object
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
