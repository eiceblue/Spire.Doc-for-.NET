using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;

namespace AppendBubbleChart
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
            section.AddParagraph().AppendText("Bubble chart.");

            // Add a new paragraph to the section
            Paragraph newPara = section.AddParagraph();

            // Append a bubble chart shape to the paragraph with specified width and height
            ShapeObject shape = newPara.AppendChart(ChartType.Bubble, 500, 300);

            // Get the chart object from the shape
            Chart chart = shape.Chart;

            // Clear any existing series in the chart
            chart.Series.Clear();

            // Add a new series to the chart with data points for X, Y, and bubble size values
            ChartSeries series = chart.Series.Add("Test Series",
                new[] { 2.9, 3.5, 1.1, 4.0, 4.0 },
                new[] { 1.9, 8.5, 2.1, 6.0, 1.5 },
                new[] { 9.0, 4.5, 2.5, 8.0, 5.0 });

            // Save the document to a file in Docx format
            document.SaveToFile("AppendBubbleChart.docx", FileFormat.Docx);

            // Dispose of the document object when finished using it
            document.Dispose();

            //Launch the Word file.
            WordDocViewer("AppendBubbleChart.docx");
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
