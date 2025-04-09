using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;

namespace AppendSurface3DChart
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
            section.AddParagraph().AppendText("Surface3D chart.");

            // Add a new paragraph to the section
            Paragraph newPara = section.AddParagraph();

            // Append a Surface3D chart shape to the paragraph with specified width and height
            ShapeObject shape = newPara.AppendChart(ChartType.Surface3D, 500, 300);

            // Get the chart object from the shape
            Chart chart = shape.Chart;

            // Clear any existing series in the chart
            chart.Series.Clear();

            // Set the title of the chart
            chart.Title.Text = "My chart";

            // Add multiple series to the chart with categories (X-axis values) and corresponding data values
            chart.Series.Add("Series 1",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
                new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

            chart.Series.Add("Series 2",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
                new double[] { 900000, 50000, 1100000, 400000, 2500000 });

            chart.Series.Add("Series 3",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
                new double[] { 500000, 820000, 1500000, 400000, 100000 });

            // Save the document to a file in Docx format
            document.SaveToFile("AppendSurface3DChart.docx", FileFormat.Docx);

            // Dispose of the document object when finished using it
            document.Dispose();

            //Launching the Word file.
            WordDocViewer("AppendSurface3DChart.docx");
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
