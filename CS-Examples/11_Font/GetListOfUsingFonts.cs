using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace GetListOfUsingFonts
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\UsingFonts.docx";
			string output = @"GetListOfUsingFonts.txt";
			StringBuilder stringBuilder = new StringBuilder();

			//Create a dictionary to store font and text range
			Dictionary<Font, TextRange> font_obj = new Dictionary<Font, TextRange>() { };

			//Create a Word document.
			Document document = new Document();

			//Load the file from disk.
			document.LoadFromFile(input);

			//Loop through the sections
			foreach (Section section in document.Sections)
			{
				//Loop through the paragraphs
				foreach (Paragraph paragraph in section.Body.Paragraphs)
				{

					//Loop through the child objects of the paragraph
					foreach (DocumentObject obj in paragraph.ChildObjects)
					{
						//Determine the Document Object Type of the child object
						if (obj.DocumentObjectType.Equals(DocumentObjectType.TextRange))
						{
							TextRange range = obj as TextRange;

							//Get the font 
							Font font = range.CharacterFormat.Font;

							// Determine if the font is already exists or not
							if (!font_obj.ContainsKey(font))
							{
								font_obj.Add(font, range);
							}

						}
					}
				}
			}


			//Loop through dictionary
			foreach (var item in font_obj)
			{
				//Get the font
				Font font = item.Key;

				//Get the text range
				TextRange range = item.Value;

				//Format the font name, size,style and color
				string s = string.Format("Font Name: {0}, Size:{1}, Style:{2}, Color:{3}", font.Name, font.Size, font.Style, range.CharacterFormat.TextColor.Name);
				stringBuilder.AppendLine(s);
			}

			//Write to a txt file
			File.WriteAllText(output, stringBuilder.ToString());

			//Dispose the Document
			document.Dispose();

            //Launching the Text file.
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
