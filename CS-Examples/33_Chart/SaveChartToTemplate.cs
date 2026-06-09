using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;
using Spire.Doc.Fields.Shapes;

namespace SaveChartToTemplate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Word document
            Document doc = new Document();

            // Add a new section to the document
            Section section = doc.AddSection();

            // Add a new paragraph to a newly created section
            Paragraph paragraph = section.AddParagraph();

            // Append a column chart to the paragraph and retrieve the Chart object
            Chart chart = ((Shape)paragraph.AppendChart(ChartType.Column, 400, 300)).Chart;

            // Save the chart as a template file (.crtx)
            chart.SaveAsTemplate("SaveChartToTemplate.crtx");

            // Close the document and release associated resources
            doc.Close();

            // Dispose of the document object to free up memory
            doc.Dispose();

            this.Close();         
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
