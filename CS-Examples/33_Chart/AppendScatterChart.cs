using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;

namespace AppendScatterChart
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
            section.AddParagraph().AppendText("Scatter chart.");

            // Add a new paragraph to the section
            Paragraph newPara = section.AddParagraph();

            // Append a scatter chart shape to the paragraph with specified width and height
            ShapeObject shape = newPara.AppendChart(ChartType.Scatter, 450, 300);
            Chart chart = shape.Chart;

            // Clear any existing series in the chart
            chart.Series.Clear();

            // Add a new series to the chart with data points for X and Y values
            chart.Series.Add("Scatter chart",
                new[] { 1.0, 2.0, 3.0, 4.0, 5.0 },
                new[] { 1.0, 20.0, 40.0, 80.0, 160.0 });

            // Save the document to a file in Docx format
            document.SaveToFile("AppendScatterChart.docx", FileFormat.Docx);

            // Dispose of the document object when finished using it
            document.Dispose();

            //Launching the Word file.
            WordDocViewer("AppendScatterChart.docx");
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
