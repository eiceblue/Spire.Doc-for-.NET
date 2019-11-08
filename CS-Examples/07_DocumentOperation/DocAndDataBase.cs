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
            String inputDataBase = @"..\..\..\..\..\..\Data\demo.mdb";
            String inputFolder = @"..\..\..\..\..\..\Data\";
            String fileName = "Template.docx";

            // Open a database connection
            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + inputDataBase;
            OleDbConnection connection = new OleDbConnection(connString);
            connection.Open();
 
            // Store the document to the database.
            StoreToDatabase(inputFolder + fileName, connection);
            // Read the document from the database and store the file to disk.
            Document dbDoc = ReadFromDatabase(fileName, connection);

            // Save the retrieved document to disk.
            string newFileName = "DocAndDataBase_out.docx";
            dbDoc.SaveToFile(newFileName, FileFormat.Docx);

            // Delete the document from the database.
            DeleteFromDatabase(fileName, connection);

            // Close the connection to the database.
            connection.Close();

            //Launching the MS Word file.
            WordDocViewer("DocAndDataBase_out.docx");

        }
        //Store document to database 
        public static void StoreToDatabase(String input, OleDbConnection connection)
        {
            Document doc=new Document(input);
            // Save the document to a MemoryStream object.
            MemoryStream stream = new MemoryStream();
            doc.SaveToStream(stream, FileFormat.Docx);

            // Get the filename from the document.
            string fileName = Path.GetFileName(input);

            // Create the SQL command.
            string commandString = "INSERT INTO Documents (FileName, FileContent) VALUES('" + fileName + "', @Doc)";
            OleDbCommand command = new OleDbCommand(commandString, connection);

            // Add the @Doc parameter.
            command.Parameters.AddWithValue("Doc", stream.ToArray());

            // Write the document to the database.
            command.ExecuteNonQuery();
        }

        // Read document from database 
        public static Document ReadFromDatabase(string fileName, OleDbConnection mConnection)
        {
            // Create the SQL command.
            string commandString = "SELECT * FROM Documents WHERE FileName='" + fileName + "'";
            OleDbCommand command = new OleDbCommand(commandString, mConnection);

            // Create the data adapter.
            OleDbDataAdapter adapter = new OleDbDataAdapter(command);

            // Fill the results from the database into a DataTable.
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);

            // Check whether there was a matching record found from the database and throw an exception if no record was found.
            if (dataTable.Rows.Count == 0)
                throw new ArgumentException(string.Format("Could not find any record matching the document \"{0}\" in the database.", fileName));

            // The document is stored in byte form in the FileContent column.
            // Retrieve these bytes of the first matching record to a new buffer.
            byte[] buffer = (byte[])dataTable.Rows[0]["FileContent"];

            // Wrap the bytes from the buffer into a new MemoryStream object.
            MemoryStream newStream = new MemoryStream(buffer);

            // Read the document from the stream.
            Document doc = new Document(newStream);

            // Return the retrieved document.
            return doc;
        }

        // Delete document from database 
        public static void DeleteFromDatabase(string fileName, OleDbConnection mConnection)
        {
            // Create the SQL command.
            string commandString = "DELETE * FROM Documents WHERE FileName='" + fileName + "'";
            OleDbCommand command = new OleDbCommand(commandString, mConnection);

            // Delete the record.
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
