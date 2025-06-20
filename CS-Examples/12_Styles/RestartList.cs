using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RestartList
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

			//Create a new section
			Section section = document.AddSection();

			//Create a new paragraph
			Paragraph paragraph = section.AddParagraph();

			//Append Text
			paragraph.AppendText("List 1");


			//Create a numberList
			ListStyle numberList = new ListStyle(document, ListType.Numbered);

			//Set the name
			numberList.Name = "Numbered1";

			//Add the numberList to document
			document.ListStyles.Add(numberList);

			//Add paragraph and apply the list style
			paragraph = section.AddParagraph();
			paragraph.AppendText("List Item 1");
			paragraph.ListFormat.ApplyStyle(numberList.Name);

			paragraph = section.AddParagraph();
			paragraph.AppendText("List Item 2");
			paragraph.ListFormat.ApplyStyle(numberList.Name);

			paragraph = section.AddParagraph();
			paragraph.AppendText("List Item 3");
			paragraph.ListFormat.ApplyStyle(numberList.Name);

			paragraph = section.AddParagraph();
			paragraph.AppendText("List Item 4");
			paragraph.ListFormat.ApplyStyle(numberList.Name);

			//Append Text
			paragraph = section.AddParagraph();
			paragraph.AppendText("List 2");

			ListStyle numberList2 = new ListStyle(document, ListType.Numbered);
			numberList2.Name = "Numbered2";
			//set start number of second list
			numberList2.Levels[0].StartAt = 10;
			document.ListStyles.Add(numberList2);

			//Add paragraph and apply the list style
			paragraph = section.AddParagraph();
			paragraph.AppendText("List Item 5");
			paragraph.ListFormat.ApplyStyle(numberList2.Name);

			paragraph = section.AddParagraph();
			paragraph.AppendText("List Item 6");
			paragraph.ListFormat.ApplyStyle(numberList2.Name);

			paragraph = section.AddParagraph();
			paragraph.AppendText("List Item 7");
			paragraph.ListFormat.ApplyStyle(numberList2.Name);

			paragraph = section.AddParagraph();
			paragraph.AppendText("List Item 8");
			paragraph.ListFormat.ApplyStyle(numberList2.Name);

			string output = "RestartList.docx";
			document.SaveToFile(output);

			//Dispose the document
			document.Dispose();
            WordDocViewer(output);
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
