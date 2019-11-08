using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;

namespace GetTablePosition
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a document
            Document document = new Document();
            //Load file
            document.LoadFromFile(@"..\..\..\..\..\..\Data\TableSample-Az.docx");
            //Get the first section
            Section section = document.Sections[0];
            //Get the first table
            Table table = section.Tables[0] as Table;

            StringBuilder stringBuidler = new StringBuilder();

            //Verify whether the table uses "Around" text wrapping or not.
            if (table.TableFormat.WrapTextAround)
            {
                RowFormat.TablePositioning positon = table.TableFormat.Positioning;

                stringBuidler.AppendLine("Horizontal:");               
                stringBuidler.AppendLine("Position:" + positon.HorizPosition +" pt");
                stringBuidler.AppendLine("Absolute Position:" + positon.HorizPositionAbs + ", Relative to:" + positon.HorizRelationTo);
                stringBuidler.AppendLine();
                stringBuidler.AppendLine("Vertical:");
                stringBuidler.AppendLine("Position:" + positon.VertPosition + " pt");
                stringBuidler.AppendLine("Absolute Position:" + positon.VertPositionAbs + ", Relative to:" + positon.VertRelationTo);
                stringBuidler.AppendLine();
                stringBuidler.AppendLine("Distance from surrounding text:");
                stringBuidler.AppendLine("Top:" + positon.DistanceFromTop + " pt, Left:" + positon.DistanceFromLeft + " pt");
                stringBuidler.AppendLine("Bottom:" + positon.DistanceFromBottom + "pt, Right:" + positon.DistanceFromRight + " pt");
            }


            String result = "GetTablePosition_out.txt";

            //Save file.
            File.WriteAllText(result, stringBuidler.ToString());
            //Launching the Word file.
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
