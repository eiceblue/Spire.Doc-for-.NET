using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;
using System.Drawing;

namespace AppendChartTitle
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

                            // Call the method to add or update the chart title
                            AppendChartTitle(chart);
                        }
                    }
                }
            }


            document.SaveToFile("AppendChartTitle.docx",FileFormat.Docx2019);

            //Dispose the document
            document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("AppendChartTitle.docx");
        }
        public void AppendChartTitle(Spire.Doc.Fields.Shapes.Charts.Chart chart)
        {
            // Get the chart's title object
            ChartTitle title = chart.Title;

            // Enable the display of the title
            title.Show = true;

            // Disable overlay so the title does not overlap with the chart area
            title.Overlay = false;

            // Set the text of the title
            title.Text = "My Chart";

            // Set font size of the title
            title.CharacterFormat.FontSize = 12;

            // Set the title text to bold
            title.CharacterFormat.Bold = true;

            // Set the text color to blue
            title.CharacterFormat.TextColor = Color.Blue;

            // Enable right-to-left text formatting (if needed for language)
            title.CharacterFormat.Bidi = true;

            // Apply italic style to the title text
            title.CharacterFormat.Italic = true;

            // Set character spacing (tracking or kerning)
            title.CharacterFormat.CharacterSpacing = 2;

            // Set underline color to red
            title.CharacterFormat.UnderlineColor = Color.Red;

            // Set underline style to double line
            title.CharacterFormat.UnderlineStyle = UnderlineStyle.Double;

            // Set font name
            title.CharacterFormat.FontName = "arial";

            // Enable all caps formatting
            title.CharacterFormat.AllCaps = true;

            // Enable shadow effect on the text
            title.CharacterFormat.IsShadow = true;

            // Set the position of the text baseline relative to normal 
            title.CharacterFormat.Position = 3;
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
