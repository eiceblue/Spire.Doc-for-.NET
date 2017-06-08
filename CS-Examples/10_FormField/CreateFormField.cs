using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CreateFormField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document document = new Document();
            Section section = document.AddSection();

            //page setup
            SetPage(section);

            //insert header and footer.
            InsertHeaderAndFooter(section);

            //add title
            AddTitle(section);

            //add form
            AddForm(section);

            //protect document, only form fields could be edited.
            document.Protect(ProtectionType.AllowOnlyFormFields, "e-iceblue");

            //Save doc file.
            document.SaveToFile("Sample.doc",FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer("Sample.doc");


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

        private void InsertHeaderAndFooter(Section section)
        {
            //insert picture and text to header
            Paragraph headerParagraph = section.HeadersFooters.Header.AddParagraph();
            DocPicture headerPicture
                = headerParagraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Header.png"));

            //header text
            TextRange text = headerParagraph.AppendText("Demo of Spire.Doc");
            text.CharacterFormat.FontName = "Arial";
            text.CharacterFormat.FontSize = 10;
            text.CharacterFormat.Italic = true;
            headerParagraph.Format.HorizontalAlignment
                = Spire.Doc.Documents.HorizontalAlignment.Right;

            //border
            headerParagraph.Format.Borders.Bottom.BorderType
                = Spire.Doc.Documents.BorderStyle.Single;
            headerParagraph.Format.Borders.Bottom.Space = 0.05F;


            //header picture layout - text wrapping
            headerPicture.TextWrappingStyle = TextWrappingStyle.Behind;

            //header picture layout - position
            headerPicture.HorizontalOrigin = HorizontalOrigin.Page;
            headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
            headerPicture.VerticalOrigin = VerticalOrigin.Page;
            headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top;

            //insert picture to footer
            Paragraph footerParagraph = section.HeadersFooters.Footer.AddParagraph();
            DocPicture footerPicture
                = footerParagraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Footer.png"));

            //footer picture layout
            footerPicture.TextWrappingStyle = TextWrappingStyle.Behind;
            footerPicture.HorizontalOrigin = HorizontalOrigin.Page;
            footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
            footerPicture.VerticalOrigin = VerticalOrigin.Page;
            footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom;

            //insert page number
            footerParagraph.AppendField("page number", FieldType.FieldPage);
            footerParagraph.AppendText(" of ");
            footerParagraph.AppendField("number of pages", FieldType.FieldNumPages);
            footerParagraph.Format.HorizontalAlignment
                = Spire.Doc.Documents.HorizontalAlignment.Right;

            //border
            footerParagraph.Format.Borders.Top.BorderType
                = Spire.Doc.Documents.BorderStyle.Single;
            footerParagraph.Format.Borders.Top.Space = 0.05F;
        }

        private void AddTitle(Section section)
        {
            Paragraph title = section.AddParagraph();
            TextRange titleText = title.AppendText("Create Your Account");
            titleText.CharacterFormat.FontSize = 18;
            titleText.CharacterFormat.FontName = "Arial";
            titleText.CharacterFormat.TextColor = Color.FromArgb(0x00, 0x71, 0xb6);
            title.Format.HorizontalAlignment
                = Spire.Doc.Documents.HorizontalAlignment.Center;
            title.Format.AfterSpacing = 8;
        }

        private void AddForm(Section section)
        {
            ParagraphStyle descriptionStyle = new ParagraphStyle(section.Document);
            descriptionStyle.Name = "description";
            descriptionStyle.CharacterFormat.FontSize = 12;
            descriptionStyle.CharacterFormat.FontName = "Arial";
            descriptionStyle.CharacterFormat.TextColor = Color.FromArgb(0x00, 0x45, 0x8e);
            section.Document.Styles.Add(descriptionStyle);

            Paragraph p1 = section.AddParagraph();
            String text1
                = "So that we can verify your identity and find your information, "
                + "please provide us with the following information. "
                + "This information will be used to create your online account. "
                + "Your information is not public, shared in anyway, or displayed on this site";
            p1.AppendText(text1);
            p1.ApplyStyle(descriptionStyle.Name);

            Paragraph p2 = section.AddParagraph();
            String text2
                = "You must provide a real email address to which we will send your password.";
            p2.AppendText(text2);
            p2.ApplyStyle(descriptionStyle.Name);
            p2.Format.AfterSpacing = 8;

            //field group label style
            ParagraphStyle formFieldGroupLabelStyle = new ParagraphStyle(section.Document);
            formFieldGroupLabelStyle.Name = "formFieldGroupLabel";
            formFieldGroupLabelStyle.ApplyBaseStyle("description");
            formFieldGroupLabelStyle.CharacterFormat.Bold = true;
            formFieldGroupLabelStyle.CharacterFormat.TextColor = Color.White;
            section.Document.Styles.Add(formFieldGroupLabelStyle);

            //field label style
            ParagraphStyle formFieldLabelStyle = new ParagraphStyle(section.Document);
            formFieldLabelStyle.Name = "formFieldLabel";
            formFieldLabelStyle.ApplyBaseStyle("description");
            formFieldLabelStyle.ParagraphFormat.HorizontalAlignment
                = Spire.Doc.Documents.HorizontalAlignment.Right;
            section.Document.Styles.Add(formFieldLabelStyle);

            //add table
            Table table = section.AddTable();

            //2 columns of per row
            table.DefaultColumnsNumber = 2;

            //default height of row is 20point
            table.DefaultRowHeight = 20;

            //load form config data
            using (Stream stream = File.OpenRead(@"..\..\..\..\..\..\Data\Form.xml"))
            {
                XPathDocument xpathDoc = new XPathDocument(stream);
                XPathNodeIterator sectionNodes = xpathDoc.CreateNavigator().Select("/form/section");
                foreach (XPathNavigator node in sectionNodes)
                {
                    //create a row for field group label, does not copy format
                    TableRow row = table.AddRow(false);
                    row.Cells[0].CellFormat.BackColor = Color.FromArgb(0x00, 0x71, 0xb6);
                    row.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                    //label of field group
                    Paragraph cellParagraph = row.Cells[0].AddParagraph();
                    cellParagraph.AppendText(node.GetAttribute("name", ""));
                    cellParagraph.ApplyStyle(formFieldGroupLabelStyle.Name);

                    XPathNodeIterator fieldNodes = node.Select("field");
                    foreach (XPathNavigator fieldNode in fieldNodes)
                    {
                        //create a row for field, does not copy format
                        TableRow fieldRow = table.AddRow(false);

                        //field label
                        fieldRow.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        Paragraph labelParagraph = fieldRow.Cells[0].AddParagraph();
                        labelParagraph.AppendText(fieldNode.GetAttribute("label", ""));
                        labelParagraph.ApplyStyle(formFieldLabelStyle.Name);

                        fieldRow.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        Paragraph fieldParagraph = fieldRow.Cells[1].AddParagraph();
                        String fieldId = fieldNode.GetAttribute("id", "");
                        switch (fieldNode.GetAttribute("type", ""))
                        {
                            case "text":
                                //add text input field
                                TextFormField field
                                    = fieldParagraph.AppendField(fieldId, FieldType.FieldFormTextInput) as TextFormField;

                                //set default text
                                field.DefaultText = "";
                                field.Text = "";
                                break;

                            case "list":
                                //add dropdown field
                                DropDownFormField list
                                    = fieldParagraph.AppendField(fieldId, FieldType.FieldFormDropDown) as DropDownFormField;

                                //add items into dropdown.
                                XPathNodeIterator itemNodes = fieldNode.Select("item");
                                foreach (XPathNavigator itemNode in itemNodes)
                                {
                                    list.DropDownItems.Add(itemNode.SelectSingleNode("text()").Value);
                                }
                                break;

                            case "checkbox":
                                //add checkbox field
                                fieldParagraph.AppendField(fieldId, FieldType.FieldFormCheckBox);
                                break;
                        }
                    }

                    //merge field group row. 2 columns to 1 column
                    table.ApplyHorizontalMerge(row.GetRowIndex(), 0, 1);
                }
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
