using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;

namespace CombinedChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize a new Document object
            Document doc = new Document();

            // Add a new section to the document and create a paragraph within that section
            Paragraph paragraph = doc.AddSection().AddParagraph();

            // Append a column chart (450x300 pixels) to the paragraph and retrieve the Chart object
            Chart chart = paragraph.AppendChart(ChartType.Column, 450, 300).Chart;

            // Change the chart type of the series named "Series 3" to a Line chart and enable secondary axis if applicable
            chart.ChangeSeriesType("Series 3", ChartSeriesType.Line, true);

            // Define the output file name for the combined chart document
            String outputFile = "CombinedChart.docx";

            // Save the document to the specified file in DOCX 2019 format
            doc.SaveToFile(outputFile, FileFormat.Docx2019);

            // Close the document to release resources
            doc.Close();

            // Dispose of the document object to free up memory
            doc.Dispose();

            //Launch the Word file.
            WordDocViewer(outputFile);
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
