using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Xml.XPath;

namespace FillFormField
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
			Document document = new Document(@"..\..\..\..\..\..\Data\FillFormField.doc");

			// Open the XML file containing user data
			using (Stream stream = File.OpenRead(@"..\..\..\..\..\..\Data\User.xml"))
			{
				// Create an XPathDocument from the XML stream
				XPathDocument xpathDoc = new XPathDocument(stream);

				// Select the "user" node from the XML document
				XPathNavigator user = xpathDoc.CreateNavigator().SelectSingleNode("/user");

				// Iterate through each form field in the document's first section
				foreach (FormField field in document.Sections[0].Body.FormFields)
				{
					// Get the XPath to retrieve the value for the current form field
					String path = String.Format("{0}/text()", field.Name);

					// Select the corresponding node from the XML document
					XPathNavigator propertyNode = user.SelectSingleNode(path);

					// If the node exists, set the value of the form field based on its type
					if (propertyNode != null)
					{
						switch (field.Type)
						{
							// Text input field
							case FieldType.FieldFormTextInput:
								field.Text = propertyNode.Value;
								break;

							// Dropdown field
							case FieldType.FieldFormDropDown:
								DropDownFormField combox = field as DropDownFormField;
								for (int i = 0; i < combox.DropDownItems.Count; i++)
								{
									if (combox.DropDownItems[i].Text == propertyNode.Value)
									{
										combox.DropDownSelectedIndex = i;
										break;
									}
									if (field.Name == "country" && combox.DropDownItems[i].Text == "Others")
									{
										combox.DropDownSelectedIndex = i;
									}
								}
								break;

							// Checkbox field
							case FieldType.FieldFormCheckBox:
								if (Convert.ToBoolean(propertyNode.Value))
								{
									CheckBoxFormField checkBox = field as CheckBoxFormField;
									checkBox.Checked = true;
								}
								break;
						}
					}
				}
			}

			// Save the modified document to a file
			document.SaveToFile("Sample.doc", FileFormat.Doc);

			// Dispose the document object
			document.Dispose();

            //Launch the Word file.
            WordDocViewer("Sample.doc");
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
