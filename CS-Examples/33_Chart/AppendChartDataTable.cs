using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;
using System.Drawing;

namespace AppendChartDataTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create word document
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ChartTemplate.docx");

            // Loop through all sections in the document
            for (int i = 0; i < document.Sections.Count; i++)
            {
                // Loop through all paragraphs in the current section
                for (int j = 0; j < document.Sections[i].Paragraphs.Count; j++)
                {
                    // Get the current paragraph
                    var paragraph = document.Sections[i].Paragraphs[j];

                    // Loop through all child objects in the paragraph
                    foreach (DocumentObject obj in paragraph.ChildObjects)
                    {
                        // Check if the object is a shape (e.g., chart, etc.)
                        if (obj is ShapeObject)
                        {
                            // Cast the object to a ShapeObject
                            var shape = obj as ShapeObject;

                            // Get the chart from the shape
                            Chart chart = shape.Chart;

                            // Call the method to add or update the chart data table
                            AppendChartDataTable(chart);
                        }
                    }
                }
            }


            document.SaveToFile("AppendChartDataTable.docx",FileFormat.Docx2019);

            //Dispose the document
            document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("AppendChartDataTable.docx");
        }
        public void AppendChartDataTable(Spire.Doc.Fields.Shapes.Charts.Chart chart)
        {
            // Enable the display of the data table in the chart
            chart.DataTable.Show = true;

            // Show legend keys (symbols) in the data table
            chart.DataTable.ShowLegendKeys = true;

            // Display horizontal borders between rows in the data table
            chart.DataTable.ShowHorizontalBorder = true;

            // Display vertical borders between columns in the data table
            chart.DataTable.ShowVerticalBorder = true;

            // Show an outline border around the entire data table
            chart.DataTable.ShowOutlineBorder = true;
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
