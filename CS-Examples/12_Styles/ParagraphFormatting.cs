using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

namespace Indent
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Initialize a document
            Document document = new Document();
            Section sec = document.AddSection();
            Paragraph para = sec.AddParagraph();
            para.AppendText("Paragraph Formatting");
            para.ApplyStyle(BuiltinStyle.Title);

            para = sec.AddParagraph();
            para.AppendText("This paragraph is surrounded with borders.");
            para.Format.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;
            para.Format.Borders.Color = Color.Red;

            para = sec.AddParagraph();
            para.AppendText("The alignment of this paragraph is Left.");
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

            para = sec.AddParagraph();
            para.AppendText("The alignment of this paragraph is Center.");
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            para = sec.AddParagraph();
            para.AppendText("The alignment of this paragraph is Right.");
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

            para = sec.AddParagraph();
            para.AppendText("The alignment of this paragraph is justified.");
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Justify;

            para = sec.AddParagraph();
            para.AppendText("The alignment of this paragraph is distributed.");
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Distribute;
            
            para = sec.AddParagraph();
            para.AppendText("This paragraph has the gray shadow.");
            para.Format.BackColor = Color.Gray;

            para = sec.AddParagraph();
            para.AppendText("This paragraph has the following indentations: Left indentation is 10pt, right indentation is 10pt, first line indentation is 15pt.");
            para.Format.SetLeftIndent(10);
            para.Format.SetRightIndent(10);
            para.Format.SetFirstLineIndent(15);

            para = sec.AddParagraph();
            para.AppendText("The hanging indentation of this paragraph is 15pt.");
            //Negative value represents hanging indentation
            para.Format.SetFirstLineIndent(-15);

            para = sec.AddParagraph();
            para.AppendText("This paragraph has the following spacing: spacing before is 10pt, spacing after is 20pt, line spacing is at least 10pt.");
            para.Format.AfterSpacing = 20;
            para.Format.BeforeSpacing = 10;
            para.Format.LineSpacingRule = LineSpacingRule.AtLeast;
            para.Format.LineSpacing = 10;

            //Save as docx file.
            string filePath = "ParagraphFormatting.docx";
            document.SaveToFile(filePath, FileFormat.Docx);

            //Launching the MS Word file.
            WordDocViewer(filePath);
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
