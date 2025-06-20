using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace UpdateLastSavedDate
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
			string inputFile = "../../../../../../Data/Template.docx";

			// Specify the output file path
			string resultFile = "UpdateLastSavedDate_out.docx";

			// Create a new Document object
			Document document = new Document();

			// Load the Word document from the input file
			document.LoadFromFile(inputFile);

			// Set the LastSaveDate property of the built-in document properties to the current local time converted to Greenwich time
			document.BuiltinDocumentProperties.LastSaveDate = LocalTimeToGreenwishTime(DateTime.Now);

			// Save the modified document to the output file in Docx format
			document.SaveToFile(resultFile, FileFormat.Docx);

			// Dispose the Document object
			document.Dispose();
			
            WordDocViewer(resultFile);
            
        }

        // Convert local time to Greenwich Mean Time (GMT)
		public static DateTime LocalTimeToGreenwishTime(DateTime localTime)
		{
			// Get the current local time zone
			TimeZone localTimeZone = TimeZone.CurrentTimeZone;

			// Get the time difference between local time and UTC (Coordinated Universal Time)
			TimeSpan timeSpan = localTimeZone.GetUtcOffset(localTime);

			// Subtract the time difference from the local time to get the Greenwich Mean Time (GMT)
			DateTime greenwishTime = localTime - timeSpan;

			// Return the calculated Greenwich Mean Time (GMT)
			return greenwishTime;
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
