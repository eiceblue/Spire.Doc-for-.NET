using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;
using Spire.Doc.Fields;
using System.Data.OleDb;
using System.Data;
namespace DocAndDataBase
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			// Define the input database file path
            String inputDataBase = @"..\..\..\..\..\..\Data\demo.mdb";
			
			// Define the input folder path
            String inputFolder = @"..\..\..\..\..\..\Data\";
			
			// Define the file name of the document to be stored in the database
            String fileName = "Template.docx";

            // Create a connection string using the OLE DB provider for Microsoft Access
			string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + inputDataBase;

			// Create a new OleDbConnection object with the connection string and open the connection
			OleDbConnection connection = new OleDbConnection(connString);
			connection.Open();

			// Store the document from the specified input folder in the database
			StoreToDatabase(inputFolder + fileName, connection);

			// Read the document from the database and create a Document object
			Document dbDoc = ReadFromDatabase(fileName, connection);

			// Specify the file name for the output document
			string newFileName = "DocAndDataBase_out.docx";

			// Save the document retrieved from the database to the specified file path in Docx format
			dbDoc.SaveToFile(newFileName, FileFormat.Docx);

			// Delete the document from the database
			DeleteFromDatabase(fileName, connection);

			// Close and dispose of the connection and Document objects
			connection.Close();
			connection.Dispose();
			dbDoc.Dispose();

            //Launching the MS Word file.
            WordDocViewer("DocAndDataBase_out.docx");

        }
        // Implementation of the StoreToDatabase method
		public static void StoreToDatabase(String input, OleDbConnection connection)
		{
			// Create a Document object from the specified input file
			Document doc = new Document(input);

		// Create a MemoryStream to store the document content
		MemoryStream stream = new MemoryStream();

		// Save the document to the MemoryStream in Docx format
		doc.SaveToStream(stream, FileFormat.Docx);

		// Get the file name from the input path
		string fileName = Path.GetFileName(input);

		// Define the SQL command string to insert the document into the database
		string commandString = "INSERT INTO Documents (FileName, FileContent) VALUES('" + fileName + "', @Doc)";

		// Create an OleDbCommand object with the command string and connection
		OleDbCommand command = new OleDbCommand(commandString, connection);

		// Set the parameter value for the document content using the MemoryStream
		command.Parameters.AddWithValue("Doc", stream.ToArray());

		// Execute the SQL command to store the document in the database
		command.ExecuteNonQuery();
		}

		// Implementation of the ReadFromDatabase method
		public static Document ReadFromDatabase(string fileName, OleDbConnection mConnection)
		{
			// Define the SQL command string to select the document from the database
			string commandString = "SELECT * FROM Documents WHERE FileName='" + fileName + "'";

		// Create an OleDbCommand object with the command string and connection
		OleDbCommand command = new OleDbCommand(commandString, mConnection);

		// Create an OleDbDataAdapter object with the command
		OleDbDataAdapter adapter = new OleDbDataAdapter(command);

		// Create a DataTable to store the retrieved data
		DataTable dataTable = new DataTable();

		// Fill the DataTable with the data from the database
		adapter.Fill(dataTable);

		// Check if any record matching the document is found in the DataTable
		if (dataTable.Rows.Count == 0)
			throw new ArgumentException(string.Format("Could not find any record matching the document \"{0}\" in the database.", fileName));

		// Get the file content from the first row of the DataTable
		byte[] buffer = (byte[])dataTable.Rows[0]["FileContent"];

		// Create a MemoryStream from the file content
		MemoryStream newStream = new MemoryStream(buffer);

		// Create a Document object from the MemoryStream
		Document doc = new Document(newStream);

		// Return the Document object
		return doc;
		}

		// Implementation of the DeleteFromDatabase method
		public static void DeleteFromDatabase(string fileName, OleDbConnection mConnection)
		{
			// Define the SQL command string to delete the document from the database
			string commandString = "DELETE * FROM Documents WHERE FileName='" + fileName + "'";

		// Create an OleDbCommand object with the command string and connection
		OleDbCommand command = new OleDbCommand(commandString, mConnection);

		// Execute the SQL command to delete the document from the database
		command.ExecuteNonQuery();
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
