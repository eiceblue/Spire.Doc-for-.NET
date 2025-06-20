using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ComboBoxItem
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
            // Specify the input file path
            string input = @"..\..\..\..\..\..\Data\ComboBox.docx";

            // Create a new document object
            Document doc = new Document();

            // Load the document from the specified file path
            doc.LoadFromFile(input);

            // Iterate through each section in the document
            foreach (Section section in doc.Sections)
            {
                // Iterate through each document object in the section's body
                foreach (DocumentObject bodyObj in section.Body.ChildObjects)
                {
                    // Check if the document object is a StructureDocumentTag
                    if (bodyObj.DocumentObjectType == DocumentObjectType.StructureDocumentTag)
                    {
                        // Check if the StructureDocumentTag is of type ComboBox
                        if ((bodyObj as StructureDocumentTag).SDTProperties.SDTType == SdtType.ComboBox)
                        {
                            // Access the ComboBox control properties
                            SdtComboBox combo = (bodyObj as StructureDocumentTag).SDTProperties.ControlProperties as SdtComboBox;

                            // Remove an item from the ComboBox
                            combo.ListItems.RemoveAt(1);

                            // Create a new SdtListItem and add it to the ComboBox
                            SdtListItem item = new SdtListItem("D", "D");
                            combo.ListItems.Add(item);

                            // Set the selected value of the ComboBox based on the item value "D"
                            foreach (SdtListItem sdtItem in combo.ListItems)
                            {
                                if (string.CompareOrdinal(sdtItem.Value, "D") == 0)
                                {
                                    combo.ListItems.SelectedValue = sdtItem;
                                }
                            }
                        }
                    }
                }
            }

            // Specify the output file name
            string output = "ComboBoxItem.docx";

            // Save the modified document to a file in Docx 2013 format
            doc.SaveToFile(output, FileFormat.Docx2013);

            // Dispose the document object
            doc.Dispose();
			
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
    }
}
