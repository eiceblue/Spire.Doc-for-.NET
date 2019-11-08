using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertBreak
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

            Section section = document.AddSection();

            //page setup
            SetPage(section);

            //Add cover.
            InsertCover(section);

            //insert a break code
            section = document.AddSection();
            section.AddParagraph().InsertSectionBreak(SectionBreakType.NewPage);

            //add content
            InsertContent(section);

            //Save as doc file.
            document.SaveToFile("Sample.docx", FileFormat.Docx);

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");
        }

        private void SetPage(Section section)
        {
            //the unit of all measures below is point, 1point = 0.3528 mm
            section.PageSetup.PageSize = PageSize.A4;
            section.PageSetup.Margins.Top = 72f;
            section.PageSetup.Margins.Bottom = 72f;
            section.PageSetup.Margins.Left = 89.85f;
            section.PageSetup.Margins.Right = 89.85f;
        }

        private void InsertCover(Section section)
        {
            ParagraphStyle small = new ParagraphStyle(section.Document);
            small.Name = "small";
            small.CharacterFormat.FontName = "Arial";
            small.CharacterFormat.FontSize = 9;
            small.CharacterFormat.TextColor = Color.Gray;
            section.Document.Styles.Add(small);

            Paragraph paragraph = section.AddParagraph();
            paragraph.AppendText("The sample demonstrates how to insert section break.");
            paragraph.ApplyStyle(small.Name);

            Paragraph title = section.AddParagraph();
            TextRange text = title.AppendText("Field Types Supported by Spire.Doc");
            text.CharacterFormat.FontName = "Arial";
            text.CharacterFormat.FontSize = 20;
            text.CharacterFormat.Bold = true;
            title.Format.BeforeSpacing
                = section.PageSetup.PageSize.Height / 2 - 3 * section.PageSetup.Margins.Top;
            title.Format.AfterSpacing = 8;
            title.Format.HorizontalAlignment
                = Spire.Doc.Documents.HorizontalAlignment.Right;

            paragraph = section.AddParagraph();
            paragraph.AppendText("e-iceblue Spire.Doc team.");
            paragraph.ApplyStyle(small.Name);
            paragraph.Format.HorizontalAlignment
                = Spire.Doc.Documents.HorizontalAlignment.Right;
        }

        private void InsertContent(Section section)
        {
            ParagraphStyle list = new ParagraphStyle(section.Document);
            list.Name = "list";
            list.CharacterFormat.FontName = "Arial";
            list.CharacterFormat.FontSize = 11;
            list.ParagraphFormat.LineSpacing = 1.5F * 12F;
            list.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            section.Document.Styles.Add(list);

            Paragraph title = section.AddParagraph();
            TextRange text = title.AppendText("Field type list:");
            title.ApplyStyle(list.Name);

            bool first = true;
            foreach (FieldType type in Enum.GetValues(typeof(FieldType)))
            {
                if (type == FieldType.FieldUnknown
                    || type == FieldType.FieldNone || type == FieldType.FieldEmpty)
                {
                    continue;
                }
                Paragraph paragraph = section.AddParagraph();
                paragraph.AppendText(String.Format("{0} is supported in Spire.Doc", type));

                if (first)
                {
                    paragraph.ListFormat.ApplyNumberedStyle();
                    first = false;
                }
                else
                {
                    paragraph.ListFormat.ContinueListNumbering();
                }
                paragraph.ApplyStyle(list.Name);
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
