using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;

namespace AppendBarChart
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
            section.AddParagraph().AppendText("Bar chart.");

            // Add a new paragraph to the section
            Paragraph newPara = section.AddParagraph();

            // Append a bar chart shape to the paragraph with specified width and height
            ShapeObject chartShape = newPara.AppendChart(ChartType.Bar, 400, 300);
            Chart chart = chartShape.Chart;

            // Get the title of the chart
            ChartTitle title = chart.Title;

            // Set the text of the chart title
            title.Text = "My Chart";

            // Show the chart title
            title.Show = true;

            // Overlay the chart title on top of the chart
            title.Overlay = true;

            // Save the document to a file in Docx format
            document.SaveToFile("AppendBarChart.docx", FileFormat.Docx);

            // Dispose of the document object when finished using it
            document.Dispose();

            //Launch the Word file.
            WordDocViewer("AppendBarChart.docx");

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
