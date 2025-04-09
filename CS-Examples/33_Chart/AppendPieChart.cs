using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;

namespace AppendPieChart
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
            section.AddParagraph().AppendText("Pie chart.");

            // Add a new paragraph to the section
            Paragraph newPara = section.AddParagraph();

            // Append a pie chart shape to the paragraph with specified width and height
            ShapeObject shape = newPara.AppendChart(ChartType.Pie, 500, 300);
            Chart chart = shape.Chart;

            // Add a series to the chart with categories (labels) and corresponding data values
            ChartSeries series = chart.Series.Add("Test Series",
                new[] { "Word", "PDF", "Excel" },
                new[] { 2.7, 3.2, 0.8 });

            // Save the document to a file in Docx format
            document.SaveToFile("AppendPieChart.docx", FileFormat.Docx);

            // Dispose of the document object when finished using it
            document.Dispose();

            //Launch the Word file.
            WordDocViewer("AppendPieChart.docx");
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
