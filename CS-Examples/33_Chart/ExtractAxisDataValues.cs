using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;
using System.Text;
using System.IO;

namespace ExtractAxisDataValues
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object for the first document
            Document doc = new Document();

            // Load the Word document from the specified relative file path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractAxisDataValues.docx");

            // Initialize a StringBuilder to store the extracted axis data values
            StringBuilder stringBuilder = new StringBuilder();

            // Loop through each section in the loaded document
            foreach (Section sec in doc.Sections)
            {
                // Loop through each paragraph within the current section
                foreach (Paragraph paragraph in sec.Paragraphs)
                {
                    // Iterate over all child objects contained in the paragraph
                    for (int i = 0; i < paragraph.ChildObjects.Count; i++)
                    {
                        // Get the current document object at index i
                        DocumentObject obj = paragraph.ChildObjects[i];

                        // Check if the current object is a ShapeObject (which can contain charts)
                        if (obj is ShapeObject)
                        {
                            // Cast the object to a ShapeObject
                            ShapeObject shape = obj as ShapeObject;

                            // Retrieve the Chart object from the shape
                            Chart chart = shape.Chart;

                            // Add a header line indicating the start of X-axis data extraction
                            stringBuilder.AppendLine("Obtain X-axis data values:");

                            // Loop through all the X-axis values in the chart
                            for (int x = 0; x < chart.XValues.Count; x++)
                            {
                                // Get the specific X-axis value at the current index
                                ChartValue xVal = chart.XValues[x];

                                // Append the string representation of the X-value to the StringBuilder, followed by a space
                                stringBuilder.Append(xVal.StringValue + " ");
                            }

                            // Get the first data series from the chart (index 0)
                            ChartSeries series = chart.Series[0];

                            // Add a new line and a header for Y-axis data extraction
                            stringBuilder.AppendLine("rnObtain Y-axis data values:");

                            // Iterate through all the Y-values in the selected data series
                            foreach (ChartValue yVal in series.YValues)
                            {
                                // Append the numeric value of the Y-data point to the StringBuilder, followed by a space
                                stringBuilder.Append(yVal.Value + " ");
                            }
                        }
                    }
                }
            }

            // Define the output file name/path for saving the extracted data
            String result = "ExtractAxisDataValues.txt";

            // Write the entire content of the StringBuilder to the specified text file
            File.WriteAllText(result, stringBuilder.ToString());

            // Close the document and release associated resources
            doc.Close();

            // Dispose of the document object to free up memory
            doc.Dispose();

            WordDocViewer(result);
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
