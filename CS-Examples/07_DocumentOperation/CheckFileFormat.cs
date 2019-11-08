using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;

namespace CheckFileFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
               String input = @"..\..\..\..\..\..\Data\Template.docx";
               
               Document doc = new Document();
               doc.LoadFromFile(input);
              //Get file format
               FileFormat ff = doc.DetectedFormatType;
               string fileFormat ="The file format is ";
               //Check the format info
               switch (ff)
               {
                   case FileFormat.Doc:
                       fileFormat += "Microsoft Word 97-2003 document.";
                       break;
                   case FileFormat.Dot:
                       fileFormat += "Microsoft Word 97-2003 template.";
                       break;
                   case FileFormat.Docx:
                       fileFormat += "Office Open XML WordprocessingML Macro-Free Document.";
                       break;
                   case FileFormat.Docm:
                       fileFormat += "Office Open XML WordprocessingML Macro-Enabled Document.";
                       break;
                   case FileFormat.Dotx:
                       fileFormat += "Office Open XML WordprocessingML Macro-Free Template.";
                       break;
                   case FileFormat.Dotm:
                       fileFormat += "Office Open XML WordprocessingML Macro-Enabled Template.";
                       break;
                   case FileFormat.Rtf:
                       fileFormat += "RTF format.";
                       break;
                   case FileFormat.WordML:
                       fileFormat += "Microsoft Word 2003 WordprocessingML format.";
                       break;
                   case FileFormat.Html:
                       fileFormat += "HTML format.";
                       break;
                   case FileFormat.WordXml:
                       fileFormat += "Microsoft Word xml format for word 2007-2013.";
                       break;
                   case FileFormat.Odt:
                       fileFormat += "OpenDocument Text.";
                       break;
                   case FileFormat.Ott:
                       fileFormat += "OpenDocument Text Template.";
                       break;
                   case FileFormat.DocPre97:
                       fileFormat += "Microsoft Word 6 or Word 95 format.";
                       break;
                   default:
                        fileFormat +="Unknown format.";
                       break;
               }
               MessageBox.Show(fileFormat);
        }
    }
}
