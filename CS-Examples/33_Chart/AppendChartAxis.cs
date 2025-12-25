using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Fields;
using System.Drawing;

namespace AppendChartAxis
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

                            // Call the method to add or update the chart axis
                            AppendChartAxis(chart);
                        }
                    }
                }
            }


            document.SaveToFile("AppendChartAxis.docx", FileFormat.Docx2019);

            //Dispose the document
            document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("AppendChartAxis.docx");
        }
        public void AppendChartAxis(Spire.Doc.Fields.Shapes.Charts.Chart chart)
        {
            for (int i = 0; i < chart.Axes.Count; i++)
            {
                if (i == 0)
                {
                    chart.Axes[i].CategoryType = AxisCategoryType.Category;
                    chart.Axes[i].Bounds.Maximum = new AxisBound(5);
                    chart.Axes[i].Bounds.Minimum = new AxisBound(0);
                    chart.Axes[i].Units.Major = 1;
                    chart.Axes[i].Units.MajorTimeUnit = 0;
                    chart.Axes[i].Units.Minor = 1;
                    chart.Axes[i].Units.MinorTimeUnit = AxisTimeUnit.Days;
                    chart.Axes[i].HasMajorGridlines = false;
                    chart.Axes[i].HasMinorGridlines = true;
                    chart.Axes[i].Labels.IsAutoSpacing = false;
                    chart.Axes[i].Labels.Spacing = 1;
                    chart.Axes[i].Labels.Offset = 1;
                    chart.Axes[i].Labels.Position = AxisTickLabelPosition.Low;
                    chart.Axes[i].ReverseOrder = true;
                    chart.Axes[i].Title.Text = "x-axis";
                    chart.Axes[i].Title.Show = true;
                    chart.Axes[i].Title.Overlay = true;
                }
                else if (i == 1)
                {
                    chart.Axes[i].CategoryType = 0;
                    chart.Axes[i].Units.IsMajorAuto = true;
                    chart.Axes[i].Units.IsMinorAuto = true;
                    chart.Axes[i].Bounds.LogBase = 10;
                    chart.Axes[i].HasMajorGridlines = true;
                    chart.Axes[i].HasMinorGridlines = false;
                    chart.Axes[i].ReverseOrder = false;
                    chart.Axes[i].Labels.IsAutoSpacing = true;
                    chart.Axes[i].Title.Text = "y-axis";
                    chart.Axes[i].Title.Show = true;
                    chart.Axes[i].Title.Overlay = true;
                }
                else
                {
                    chart.Axes[i].Title.Text = "z-axis";
                    chart.Axes[i].Title.Show = true;
                    chart.Axes[i].Title.Overlay = false;
                }
                chart.Axes[i].Labels.Alignment = LabelAlignment.Left; 
                chart.Axes[i].Units.BaseTimeUnit = 0;
                chart.Axes[i].AxisBetweenCategories = true; 
                chart.Axes[i].DisplayUnits.CustomUnit = 1;
                chart.Axes[i].DisplayUnits.Unit = AxisBuiltInUnit.Custom;
                chart.Axes[i].DisplayUnits.ShowLabel = true;
                chart.Axes[i].TickMarks.Spacing = 1;
                chart.Axes[i].TickMarks.Major = 0;
                chart.Axes[i].TickMarks.Minor = AxisTickMark.Inside;
                chart.Axes[i].Title.GetCharacterFormat().FontSize = 8;
                chart.Axes[i].Title.GetCharacterFormat().TextColor = Color.Red;
                chart.Axes[i].Title.GetCharacterFormat().Bold = true;
            }
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
