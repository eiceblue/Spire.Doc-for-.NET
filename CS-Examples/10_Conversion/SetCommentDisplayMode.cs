using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Collections;
using System.Text;
using Spire.Doc.Layout;

namespace SetCommentDisplayMode
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			      // Load the document from a file
		       	Document document = new Document(@"..\..\..\..\..\..\..\Data\CommentSample.docx");

            // Set comment display mode when converting to pdf
            document.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;
            //document.LayoutOptions.CommentDisplayMode = CommentDisplayMode.Hide;
            //document.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInBalloons;


            document.SaveToFile("SetCommentDisplayMode.pdf", FileFormat.PDF);
            // Dispose the document object
            document.Dispose();

            //Launch result file
            WordDocViewer("SetCommentDisplayMode.pdf");

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
