using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;
using System.Drawing;

namespace AppendChartLegend
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create word document
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ChartTemplate.docx");

            // Loop through all sections in the document
            for (int i = 0; i < document.Sections.Count; i++)
            {
                // Loop through all paragraphs in the current section
                for (int j = 0; j < document.Sections[i].Paragraphs.Count; j++)
                {
                    // Get the current paragraph
                    var paragraph = document.Sections[i].Paragraphs[j];

                    // Loop through all child objects in the paragraph
                    foreach (DocumentObject obj in paragraph.ChildObjects)
                    {
                        // Check if the object is a shape (e.g., chart, etc.)
                        if (obj is ShapeObject)
                        {
                            // Cast the object to a ShapeObject
                            var shape = obj as ShapeObject;

                            // Get the chart from the shape
                            Chart chart = shape.Chart;

                            // Call the method to add or update the chart legend
                            AppendChartLegend(chart);
                        }
                    }
                }
            }


            document.SaveToFile("AppendChartLegend.docx", FileFormat.Docx2019);

            //Dispose the document
            document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("AppendChartLegend.docx");
        }
        public void AppendChartLegend(Spire.Doc.Fields.Shapes.Charts.Chart chart)
        {
            // Enable the legend display on the chart
            chart.Legend.Show = true;

            // Set the position of the legend to the left side of the chart
            chart.Legend.Position = LegendPosition.Left;

            // Disable overlay mode so the legend does not overlap with the chart plot area
            chart.Legend.Overlay = false;

            // Set the font size of the legend text to 9 points
            chart.Legend.CharacterFormat.FontSize = 9;

            // Set the text color of the legend labels to blue
            chart.Legend.CharacterFormat.TextColor = Color.Blue;

            // Apply italic style to the legend text
            chart.Legend.CharacterFormat.Italic = true;
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
