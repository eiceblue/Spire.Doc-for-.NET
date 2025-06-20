using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CustomizeTableOfContent
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
       
			// Create a new document
			Document doc = new Document();

			// Add a section to the document
			Section section = doc.AddSection();

			// Create a table of contents and add it to a paragraph in the section
			TableOfContent toc = new TableOfContent(doc, "{\\o \"1-3\" \\n 1-1}");
			Paragraph para = section.AddParagraph();
			para.Items.Add(toc);
			para.AppendFieldMark(FieldMarkType.FieldSeparator);
			para.AppendText("TOC");
			para.AppendFieldMark(FieldMarkType.FieldEnd);
			doc.TOC = toc;

			// Add a paragraph to the section 
			Paragraph par = section.AddParagraph();
			TextRange tr = par.AppendText("Flowers");
			tr.CharacterFormat.FontSize = 30;
			par.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			// Add a paragraph to the section 
			Paragraph para1 = section.AddParagraph();
			para1.AppendText("Ornithogalum");

			// Apply the "Heading1" style 
			para1.ApplyStyle(BuiltinStyle.Heading1);

			// Add a paragraph to the section and insert an image 
			para1 = section.AddParagraph();
			DocPicture picture = para1.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Ornithogalum.jpg"));
			picture.TextWrappingStyle = TextWrappingStyle.Square;
			para1.AppendText("Ornithogalum is a genus of perennial plants mostly native to southern Europe and southern Africa belonging to the family Asparagaceae. Some species are native to other areas such as the Caucasus. Growing from a bulb, species have linear basal leaves and a slender stalk, up to 30 cm tall, bearing clusters of typically white star-shaped flowers, often striped with green.");
			para1 = section.AddParagraph();

			// Add a paragraph to the section 
			Paragraph para2 = section.AddParagraph();
			para2.AppendText("Rosa");

			// Apply the "Heading2" style 
			para2.ApplyStyle(BuiltinStyle.Heading2);

			// Add a paragraph to the section and insert an image 
			para2 = section.AddParagraph();
			DocPicture picture2 = para2.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Rosa.jpg"));
			picture2.TextWrappingStyle = TextWrappingStyle.Square;
			para2.AppendText("A rose is a woody perennial flowering plant of the genus Rosa, in the family Rosaceae, or the flower it bears. There are over a hundred species and thousands of cultivars. They form a group of plants that can be erect shrubs, climbing or trailing with stems that are often armed with sharp prickles. Flowers vary in size and shape and are usually large and showy, in colours ranging from white through yellows and reds. Most species are native to Asia, with smaller numbers native to Europe, North America, and northwestern Africa. Species, cultivars and hybrids are all widely grown for their beauty and often are fragrant. Roses have acquired cultural significance in many societies. Rose plants range in size from compact, miniature roses, to climbers that can reach seven meters in height. Different species hybridize easily, and this has been used in the development of the wide range of garden roses.");
			section.AddParagraph();

			// Add a paragraph to the section 
			Paragraph para3 = section.AddParagraph();
			para3.AppendText("Hyacinth");

			// Apply the "Heading3" style 
			para3.ApplyStyle(BuiltinStyle.Heading3);

			// Add a paragraph to the section and insert an image 
			para3 = section.AddParagraph();
			DocPicture picture3 = para3.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\hyacinths.JPG"));
			picture3.TextWrappingStyle = TextWrappingStyle.Tight;
			para3.AppendText("Hyacinthus is a small genus of bulbous, fragrant flowering plants in the family Asparagaceae, subfamily Scilloideae. These are commonly called hyacinths. The genus is native to the eastern Mediterranean (from the south of Turkey through to northern Israel).");
			para3 = section.AddParagraph();
			para3.AppendText("Several species of Brodiea, Scilla, and other plants that were formerly classified in the lily family and have flower clusters borne along the stalk also have common names with the word \"hyacinth\" in them. Hyacinths should also not be confused with the genus Muscari, which are commonly known as grape hyacinths.");

			// Update the table of contents
			doc.UpdateTableOfContents();

		    // Save the document to a file
		    doc.SaveToFile("result.docx", FileFormat.Docx);

		    // Dispose of the document object
		    doc.Dispose();

            //Launch the Word file
            FileViewer("result.docx");
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
            this.pictureBox1.Image = global::CustomizeTableOfContent.Properties.Resources.Word;
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
            this.button1.Location = new System.Drawing.Point(367, 80);
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
            this.Text = "Customize Table Of Content";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;

    }
}
