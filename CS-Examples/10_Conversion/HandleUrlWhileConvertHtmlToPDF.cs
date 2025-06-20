using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace HandleUrlWhileConvertHtmlToPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class
            Document document = new Document();

            // Subscribe to the HtmlUrlLoadEvent to handle external resource loading.
            document.HtmlUrlLoadEvent += MyDownloadEvent;

            // Load an HTML file into the document. The file path and validation type are specified.
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_HtmlFile3.html", FileFormat.Html, XHTMLValidationType.None);

            // Save the loaded HTML content as a PDF file.
            document.SaveToFile("HtmlFileToPDF.pdf", FileFormat.PDF);

            // Dispose the document object to release resources.
            document.Dispose();
            //Launching the pdf reader to open.
            FileViewer("HtmlFileToPDF.pdf");
        }

        private static void MyDownloadEvent(object sender, Document.HtmlUrlLoadEventArgs args)
        {
            // Use WebClient to download external resources (e.g., images, CSS files) from URLs referenced in the HTML.
            using (WebClient webClient = new WebClient())
            {
                // Use the default credentials for authentication.
                webClient.Credentials = CredentialCache.DefaultCredentials;

                // Set a custom user-agent header to mimic a web browser during resource download.
                webClient.Headers.Set("user-agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0");

                // Configure supported security protocols for SSL/TLS connections.
                // This ensures compatibility with different server configurations.
                // SystemDefault = 0, Ssl3 = 48, Tls = 192, Tls11 = 768, Tls12 = 3072, Tls13 = 12288
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)0 |
                                                       (SecurityProtocolType)12288 |  // Tls13
                                                       (SecurityProtocolType)3072 |  // Tls12
                                                       (SecurityProtocolType)768 |   // Tls11
                                                       (SecurityProtocolType)192 |   // Tls
                                                       (SecurityProtocolType)48;     // Ssl3

                // Download the resource data from the provided URL.
                byte[] webData = webClient.DownloadData(args.Url);

                // Set the downloaded data into the event arguments for further processing.
                args.DataBytes = webData;
            }
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::HandleUrlWhileConvertHtmlToPDF.Properties.Resources.Word;
            this.pictureBox1.Location = new System.Drawing.Point(12, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(56, 48);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.BackColor = System.Drawing.Color.Transparent;
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(376, 80);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(96, 27);
            this.button1.TabIndex = 63;
            this.button1.Text = "Run";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Verdana", 8.25F);
            this.label1.Location = new System.Drawing.Point(85, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(387, 65);
            this.label1.TabIndex = 64;
            this.label1.Text = resources.GetString("label1.Text");
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(484, 122);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Handle Url While Convert Html To PDF";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;

    }
}
