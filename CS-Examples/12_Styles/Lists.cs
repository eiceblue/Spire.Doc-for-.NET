using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace Text
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a Word document
			Document document = new Document();

			//Add a section
			Section sec = document.AddSection();

			//Add paragraph and set list style
			Paragraph paragraph = sec.AddParagraph();

			//Append text
			paragraph.AppendText("Lists");

			//Apply the Title style
			paragraph.ApplyStyle(BuiltinStyle.Title);

			//Add a new paragraph
			paragraph = sec.AddParagraph();

			//Append text and set the the bold style 
			paragraph.AppendText("Numbered List:").CharacterFormat.Bold = true;

			//Create list style
			ListStyle numberList = new ListStyle(document, ListType.Numbered);
			numberList.Name = "numberList";
			numberList.Levels[1].NumberPrefix = "\x0000.";
			numberList.Levels[1].PatternType = ListPatternType.Arabic;
			numberList.Levels[2].NumberPrefix = "\x0000.\x0001.";
			numberList.Levels[2].PatternType = ListPatternType.Arabic;

			ListStyle bulletList = new ListStyle(document, ListType.Bulleted);
			bulletList.Name = "bulletList";

			//add the list styles into document
			document.ListStyles.Add(numberList);
			document.ListStyles.Add(bulletList);

			//Add paragraph and apply the list style
			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 1");

			//Apply the number List style
			paragraph.ListFormat.ApplyStyle(numberList.Name);

			//Add a paragrahph
			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 2");
			paragraph.ListFormat.ApplyStyle(numberList.Name);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 2.1");
			paragraph.ListFormat.ApplyStyle(numberList.Name);
			paragraph.ListFormat.ListLevelNumber = 1;

			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 2.2");
			paragraph.ListFormat.ApplyStyle(numberList.Name);
			paragraph.ListFormat.ListLevelNumber = 1;

			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 2.2.1");
			paragraph.ListFormat.ApplyStyle(numberList.Name);
			paragraph.ListFormat.ListLevelNumber = 2;
			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 2.2.2");
			paragraph.ListFormat.ApplyStyle(numberList.Name);
			paragraph.ListFormat.ListLevelNumber = 2;
			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 2.2.3");
			paragraph.ListFormat.ApplyStyle(numberList.Name);
			paragraph.ListFormat.ListLevelNumber = 2;

			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 2.3");
			paragraph.ListFormat.ApplyStyle(numberList.Name);
			paragraph.ListFormat.ListLevelNumber = 1;

			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 3");
			paragraph.ListFormat.ApplyStyle(numberList.Name);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("Bulleted List:").CharacterFormat.Bold = true;


			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 1");

			//Apply the bullet list style
			paragraph.ListFormat.ApplyStyle(bulletList.Name);


			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 2");
			paragraph.ListFormat.ApplyStyle(bulletList.Name);

			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 2.1");
			paragraph.ListFormat.ApplyStyle(bulletList.Name);
			paragraph.ListFormat.ListLevelNumber = 1;

			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 2.2");
			paragraph.ListFormat.ApplyStyle(bulletList.Name);
			paragraph.ListFormat.ListLevelNumber = 1;

			paragraph = sec.AddParagraph();
			paragraph.AppendText("List Item 3");
			paragraph.ListFormat.ApplyStyle(bulletList.Name);

			//Save doc file.
			string filePath = "Lists.docx";
			document.SaveToFile(filePath, FileFormat.Docx);

			//Dispose the document
			document.Dispose();

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
