using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Interface;

namespace AddCombinationChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document instance
            Document document = new Document();

            // Add a section to the document and then add a paragraph to that section
            Paragraph paragraph = document.AddSection().AddParagraph();

            // Append a chart of specified type and size to the paragraph, and get the Chart object
            Chart chart = paragraph.AppendChart(ChartType.Column, 450, 300).Chart;

            // Modify 'Series 3' to a line chart and set it to display on the secondary axis
            chart.ChangeSeriesType("Series 3", ChartSeriesType.Line, true);

            // Define the file path and name for saving the document
            string filePath = "AddCombinationChart.docx";

            // Save the document to the specified file path in DOCX format
            document.SaveToFile(filePath, FileFormat.Docx);

            // Release resources used by the document
            document.Dispose();

            WordDocViewer(filePath);
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
