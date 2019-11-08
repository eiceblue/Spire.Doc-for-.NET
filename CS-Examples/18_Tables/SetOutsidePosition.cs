using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetOutsidePosition
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a new word document and add new section
            Document doc = new Document();
            Section sec = doc.AddSection();

            //Get header
            HeaderFooter header = doc.Sections[0].HeadersFooters.Header;

            //Add new paragraph on header and set HorizontalAlignment of the paragraph as left
            Paragraph paragraph = header.AddParagraph();
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

            //Load an image for the paragraph
            DocPicture headerimage = paragraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Word.png"));

            //Add a table of 4 rows and 2 columns
            Table table = header.AddTable();
            table.ResetCells(4, 2);

            //Set the position of the table to the right of the image
            table.TableFormat.WrapTextAround = true;
            table.TableFormat.Positioning.HorizPositionAbs = HorizontalPosition.Outside;
            table.TableFormat.Positioning.VertRelationTo = VerticalRelation.Margin;
            table.TableFormat.Positioning.VertPosition = 43;

            //Add contents for the table
            String[][] data = {
                    new string[] {"Spire.Doc.left","Spire XLS.right"},
                    new string[] {"Spire.Presentatio.left","Spire.PDF.right"},
                    new string[] {"Spire.DataExport.left","Spire.PDFViewe.right"},
                    new string []{"Spire.DocViewer.left","Spire.BarCode.right"}
                              };

            for (int r = 0; r < 4; r++)
            {
                TableRow dataRow = table.Rows[r];
                for (int c = 0; c < 2; c++)
                {
                    if (c == 0)
                    {
                        Paragraph par = dataRow.Cells[c].AddParagraph();
                        par.AppendText(data[r][c]);
                        par.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                        dataRow.Cells[c].Width = 180;
                    }
                    else
                    {
                        Paragraph par = dataRow.Cells[c].AddParagraph();
                        par.AppendText(data[r][c]);
                        par.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                        dataRow.Cells[c].Width = 180;
                    }
                }
            }

            //Save and launch document
            string output = "SetOutsidePosition.docx";
            doc.SaveToFile(output, FileFormat.Docx);
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
