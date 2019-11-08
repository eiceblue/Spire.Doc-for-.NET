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
            // Create a new document and load from file
            string input = @"..\..\..\..\..\..\Data\ComboBox.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the combo box from the file
            foreach (Section section in doc.Sections)
            {
                foreach (DocumentObject bodyObj in section.Body.ChildObjects)
                {
                    if (bodyObj.DocumentObjectType == DocumentObjectType.StructureDocumentTag)
                    {
                        //If SDTType is ComboBox
                        if ((bodyObj as StructureDocumentTag).SDTProperties.SDTType == SdtType.ComboBox)
                        {
                            SdtComboBox combo = (bodyObj as StructureDocumentTag).SDTProperties.ControlProperties as SdtComboBox;
                            //Remove the second list item
                            combo.ListItems.RemoveAt(1);
                            //Add a new item
                            SdtListItem item = new SdtListItem("D", "D");
                            combo.ListItems.Add(item);

                            //If the value of list items is "D"
                            foreach (SdtListItem sdtItem in combo.ListItems)
                            {
                                if (string.CompareOrdinal(sdtItem.Value, "D") == 0)
                                {
                                    //Select the item
                                    combo.ListItems.SelectedValue = sdtItem;
                                }
                            }
                        }
                    }
                }
            }

            //Save the document and launch it
            string output = "ComboBoxItem.docx";
            doc.SaveToFile(output, FileFormat.Docx2013);
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
