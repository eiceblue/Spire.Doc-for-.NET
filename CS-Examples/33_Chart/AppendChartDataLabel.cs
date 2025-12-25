using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;
using System.Drawing;

namespace AppendChartDataLabel
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
                            ChartSeriesCollection series = chart.Series;
                            ChartDataLabelCollection dataLabels = series[0].DataLabels;
                            series[0].HasDataLabels = true;
                            AppendChartDataLabel(dataLabels);
                        }
                    }
                }
            }

            document.SaveToFile("AppendChartDataLabel.docx",FileFormat.Docx2019);

            //Dispose the document
            document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("AppendChartDataLabel.docx");
        }
        public void AppendChartDataLabel(ChartDataLabelCollection dataLabels)
        {
            // Display the value (e.g., percentage or numerical value) on the data labels
            dataLabels.ShowValue = true;

            // Display the category name (e.g., the label for each chart segment)
            dataLabels.ShowCategoryName = true;

            // Display the series name (useful when multiple series are present)
            dataLabels.ShowSeriesName = true;

            // Show leader lines connecting the data labels to the chart elements
            dataLabels.ShowLeaderLines = true;

            // Set the separator between different label components (e.g., value and category)
            dataLabels.Separator = ";";

            // Set the number format for the displayed values (thousands separator and zero decimals)
            dataLabels.NumberFormat.FormatCode = "#,##0";

            // Set the font size of the data labels
            dataLabels.CharacterFormat.FontSize = 8;

            // Make the text in the data labels bold
            dataLabels.CharacterFormat.Bold = true;

            // Set the text color of the data labels to blue
            dataLabels.CharacterFormat.TextColor = Color.Blue;

            // Set the border color of the characters in the data labels to blue
            dataLabels.CharacterFormat.Border.Color = Color.Blue;

            // Enable right-to-left (RTL) text direction for languages like Arabic or Hebrew
            dataLabels.CharacterFormat.Bidi = true;

            // Apply italic formatting to the text
            dataLabels.CharacterFormat.Italic = true;

            // Set the underline color to red
            dataLabels.CharacterFormat.UnderlineColor = Color.Red;

            // Set the underline style to double line
            dataLabels.CharacterFormat.UnderlineStyle = UnderlineStyle.Double;

            // Set the font family for the data labels
            dataLabels.CharacterFormat.FontName = "Arial";

            // Display all text in uppercase letters
            dataLabels.CharacterFormat.AllCaps = true;

            // Apply a shadow effect to the text
            dataLabels.CharacterFormat.IsShadow = true;

            // Set the opacity (transparency) of the text effect (e.g., shadow or glow)
            dataLabels.CharacterFormat.TextEffectFormat.TextOpacity = 0.1;
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
