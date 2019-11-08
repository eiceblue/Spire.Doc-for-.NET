using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;
using Spire.Doc.Fields;

namespace RemoveHyperlinks
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load Document
            string input = @"..\..\..\..\..\..\Data\Hyperlinks.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get all hyperlinks
            List<Field> hyperlinks = FindAllHyperlinks(doc);

            //Flatten all hyperlinks
            for (int i = hyperlinks.Count - 1; i >= 0; i--)
            {
                FlattenHyperlinks(hyperlinks[i]);
            }

            //Save and launch document
            string output = "RemoveHyperlinks.docx";
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

        //Create a method FindAllHyperlinks() to get all the hyperlinks from the sample document
        private List<Field> FindAllHyperlinks(Document document)
        {
            List<Field> hyperlinks = new List<Field>();
            //Iterate through the items in the sections to find all hyperlinks
            foreach (Section section in document.Sections)
            {
                foreach (DocumentObject sec in section.Body.ChildObjects)
                {
                    if (sec.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
                        {
                            if (para.DocumentObjectType == DocumentObjectType.Field)
                            {
                                Field field = para as Field;
                                if (field.Type == FieldType.FieldHyperlink)
                                {
                                    hyperlinks.Add(field);
                                }
                            }
                        }
                    }
                }
            }
            return hyperlinks;
        }

        // Flatten the hyperlink field
        private void FlattenHyperlinks(Field field)
        {
            int ownerParaIndex = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.OwnerParagraph);
            int fieldIndex = field.OwnerParagraph.ChildObjects.IndexOf(field);
            Paragraph sepOwnerPara = field.Separator.OwnerParagraph;
            int sepOwnerParaIndex = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.Separator.OwnerParagraph);
            int sepIndex = field.Separator.OwnerParagraph.ChildObjects.IndexOf(field.Separator);
            int endIndex = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End);
            int endOwnerParaIndex = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.End.OwnerParagraph);

            FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex);

            field.End.OwnerParagraph.ChildObjects.RemoveAt(endIndex);

            for (int i = sepOwnerParaIndex; i >= ownerParaIndex; i--)
            {
                if (i == sepOwnerParaIndex && i == ownerParaIndex)
                {
                    for (int j = sepIndex; j >= fieldIndex; j--)
                    {
                        field.OwnerParagraph.ChildObjects.RemoveAt(j);

                    }
                }
                else if (i == ownerParaIndex)
                {
                    for (int j = field.OwnerParagraph.ChildObjects.Count - 1; j >= fieldIndex; j--)
                    {
                        field.OwnerParagraph.ChildObjects.RemoveAt(j);
                    }

                }
                else if (i == sepOwnerParaIndex)
                {
                    for (int j = sepIndex; j >= 0; j--)
                    {
                        sepOwnerPara.ChildObjects.RemoveAt(j);
                    }
                }
                else
                {
                    field.OwnerParagraph.OwnerTextBody.ChildObjects.RemoveAt(i);
                }
            }
        }

        //Remove the font color and underline format of the hyperlinks
        private void FormatFieldResultText(Body ownerBody, int sepOwnerParaIndex, int endOwnerParaIndex, int sepIndex, int endIndex)
        {
            for (int i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++)
            {
                Paragraph para = ownerBody.ChildObjects[i] as Paragraph;
                if (i == sepOwnerParaIndex && i == endOwnerParaIndex)
                {
                    for (int j = sepIndex + 1; j < endIndex; j++)
                    {
                        FormatText(para.ChildObjects[j] as TextRange);
                    }

                }
                else if (i == sepOwnerParaIndex)
                {
                    for (int j = sepIndex + 1; j < para.ChildObjects.Count; j++)
                    {
                        FormatText(para.ChildObjects[j] as TextRange);
                    }
                }
                else if (i == endOwnerParaIndex)
                {
                    for (int j = 0; j < endIndex; j++)
                    {
                        FormatText(para.ChildObjects[j] as TextRange);
                    }
                }
                else
                {
                    for (int j = 0; j < para.ChildObjects.Count; j++)
                    {
                        FormatText(para.ChildObjects[j] as TextRange);
                    }
                }
            }
        }
        private void FormatText(TextRange tr)
        {
            //Set the text color to black
            tr.CharacterFormat.TextColor = Color.Black;
            //Set the text underline style to none
            tr.CharacterFormat.UnderlineStyle = UnderlineStyle.None;
        }
    }
}
