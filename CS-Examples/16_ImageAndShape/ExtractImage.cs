using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Interface;

namespace ExtractImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object
			Document document = new Document(@"..\..\..\..\..\..\Data\Template.docx");

			// Create a queue to store composite objects
			Queue<ICompositeObject> nodes = new Queue<ICompositeObject>();

			// Enqueue the document as the initial node
			nodes.Enqueue(document);

			// Create a list to store images
			IList<Image> images = new List<Image>();

			// Traverse through the composite objects in the document
			while (nodes.Count > 0)
			{
				// Dequeue the next node
				ICompositeObject node = nodes.Dequeue();

				// Iterate through the child objects of the node
				foreach (IDocumentObject child in node.ChildObjects)
				{
					// If the child is a composite object, enqueue it for further processing
					if (child is ICompositeObject)
					{
						nodes.Enqueue(child as ICompositeObject);

						// If the child is a picture, add its image to the list
						if (child.DocumentObjectType == DocumentObjectType.Picture)
						{
							DocPicture picture = child as DocPicture;
							images.Add(picture.Image);
						}
					}
				}
			}

			// Save each image in the list as a PNG file
			for (int i = 0; i < images.Count; i++)
			{
				string fileName = string.Format("Image-{0}.png", i);
				images[i].Save(fileName, ImageFormat.Png);
			}

			// If there are images, open the first one
			if (images.Count > 0)
			{
				// Open the first image using the default application
				System.Diagnostics.Process.Start("Image-0.png");
			}

			// Dispose the document
			document.Dispose();
        }
        
    }
}
