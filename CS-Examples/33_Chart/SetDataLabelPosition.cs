using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;
using Spire.Doc.Fields.Shapes;

namespace SetDataLabelPosition
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Word document instance
            Document doc = new Document();

            // Add a new section to the document
            Section section = doc.AddSection();

            // Add a paragraph with the text "Center" as a title/label
            section.AddParagraph().AppendText("Center");

            // Add a new paragraph to hold the first chart
            Spire.Doc.Documents.Paragraph newPara = section.AddParagraph();

            // Append a Pie chart to the paragraph and set its size (width: 500, height: 300)
            ShapeObject shape = newPara.AppendChart(ChartType.Pie, 500, 300);

            // Get the Chart object from the created shape
            Chart chart = shape.Chart;

            // Enable data labels for the first data series in the pie chart
            chart.Series[0].HasDataLabels = true;

            // Configure the data labels to display the category name
            chart.Series[0].DataLabels.ShowCategoryName = true;

            // Configure the data labels to display the numeric value
            chart.Series[0].DataLabels.ShowValue = true;

            // Set the position of the data labels to the center of the pie slices
            chart.Series[0].DataLabels.Position = ChartDataLabelPosition.Center;

            // Add another paragraph with the text "Left" as a title/label
            section.AddParagraph().AppendText("Left");

            newPara = section.AddParagraph();

            // Append a Bubble chart to the same paragraph and set its size (width: 500, height: 300)
            ShapeObject shape2 = newPara.AppendChart(ChartType.Bubble, 500, 300);

            // Get the Chart object from the second shape
            Chart chart2 = shape2.Chart;

            // Enable data labels for the first data series in the bubble chart
            chart2.Series[0].HasDataLabels = true;

            // Configure the data labels to display the category name
            chart2.Series[0].DataLabels.ShowCategoryName = true;

            // Configure the data labels to display the numeric value
            chart2.Series[0].DataLabels.ShowValue = true;

            // Set the position of the data labels to the left side
            chart2.Series[0].DataLabels.Position = ChartDataLabelPosition.Left;


            // Define the output file name for saving the document
            string outputFile = "SetDataLabelPosition.docx";

            // Save the document to the specified file in Docx format
            doc.SaveToFile(outputFile, FileFormat.Docx);

            // Close the document and release associated resources
            doc.Close();

            // Dispose of the document object to free up memory
            doc.Dispose();

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
