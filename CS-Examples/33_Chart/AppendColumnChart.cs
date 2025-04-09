using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;

namespace AppendColumnChart
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
            section.AddParagraph().AppendText("Column chart.");

            // Add a new paragraph to the section
            Paragraph newPara = section.AddParagraph();

            // Append a column chart shape to the paragraph with specified width and height
            ShapeObject shape = newPara.AppendChart(ChartType.Column, 500, 300);

            // Get the chart object from the shape
            Chart chart = shape.Chart;

            // Clear any existing series in the chart
            chart.Series.Clear();

            // Add a new series to the chart with data points for X values (categories) and Y values
            chart.Series.Add("Test Series",
                new[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
                new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

            // Set the number format for the Y-axis labels
            chart.AxisY.NumberFormat.FormatCode = "#,##0";

            // Save the document to a file in Docx format
            document.SaveToFile("AppendColumnChart.docx", FileFormat.Docx);

            // Dispose of the document object when finished using it
            document.Dispose();

            //Launching the Word file.
            WordDocViewer("AppendColumnChart.docx");  
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
