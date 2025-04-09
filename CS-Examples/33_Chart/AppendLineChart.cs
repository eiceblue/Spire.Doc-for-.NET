using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;

namespace AppendLineChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of Document
            Document document = new Document();

            // Add a section to the document
            Section section = document.AddSection();

            // Add a paragraph to the section and append text to it
            section.AddParagraph().AppendText("Line chart.");

            // Add a new paragraph to the section
            Paragraph newPara = section.AddParagraph();

            // Append a line chart shape to the paragraph with specified width and height
            ShapeObject shape = newPara.AppendChart(ChartType.Line, 500, 300);

            // Get the chart object from the shape
            Chart chart = shape.Chart;

            // Get the title of the chart
            ChartTitle title = chart.Title;

            // Set the text of the chart title
            title.Text = "My Chart";

            // Clear any existing series in the chart
            ChartSeriesCollection seriesColl = chart.Series;
            seriesColl.Clear();

            // Define categories (X-axis values)
            string[] categories = { "C1", "C2", "C3", "C4", "C5", "C6" };

            // Add two series to the chart with specified categories and Y-axis values
            seriesColl.Add("AW Series 1", categories, new double[] { 1, 2, 2.5, 4, 5, 6 });
            seriesColl.Add("AW Series 2", categories, new double[] { 2, 3, 3.5, 6, 6.5, 7 });

            // Save the document to a file in Docx format
            document.SaveToFile("AppendLineChart.docx", FileFormat.Docx);

            // Dispose of the document object when finished using it
            document.Dispose();

            //Launching the Word file.
            WordDocViewer("AppendLineChart.docx");
    
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
