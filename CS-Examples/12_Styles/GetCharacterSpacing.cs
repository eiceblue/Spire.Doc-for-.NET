using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace GetCharacterSpacing
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

            //Load the document from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Insert.docx");

            //Get the first section of document
            Section section = document.Sections[0];

            //Get the first paragraph 
            Paragraph paragraph = section.Paragraphs[0];

            //Define two variables
            String fontName = "";
            float fontSpacing = 0;

            //Traverse the ChildObjects 
            foreach (DocumentObject docObj in paragraph.ChildObjects)
            {
                //If it is TextRange
                if (docObj is TextRange)
                {
                    TextRange textRange = docObj as TextRange;

                    //Get the font name
                    fontName = textRange.CharacterFormat.Font.Name;

                    //Get the character spacing
                    fontSpacing = textRange.CharacterFormat.CharacterSpacing;
                }
            }

            //Show the result in message box
            MessageBox.Show("The font of first paragraph is " + fontName + ", the character spacing is " + fontSpacing + "pt.");
        }
    }
}
