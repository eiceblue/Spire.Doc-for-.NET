using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Interface;

namespace Traversal
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //open document
            Document document = new Document(@"..\..\..\..\..\..\Data\Summary_of_Science.doc");

            //document elements, each of them has child elements
            Queue<ICompositeObject> nodes = new Queue<ICompositeObject>();
            nodes.Enqueue(document);

            //embedded images list.
            IList<Image> images = new List<Image>();

            //traverse
            while (nodes.Count > 0)
            {
                ICompositeObject node = nodes.Dequeue();
                foreach (IDocumentObject child in node.ChildObjects)
                {
                    if (child is ICompositeObject)
                    {
                        nodes.Enqueue(child as ICompositeObject);
                    }
                    else if (child.DocumentObjectType == DocumentObjectType.Picture)
                    {
                        DocPicture picture = child as DocPicture;
                        images.Add(picture.Image);
                    }
                }
            }

            //save images
            for (int i = 0; i < images.Count; i++)
            {
                String fileName = String.Format("Image-{0}.png", i);
                images[i].Save(fileName, ImageFormat.Png);
            }

            if (images.Count > 0)
            {
                //show the first image
                System.Diagnostics.Process.Start("Image-0.png");
            }
        }
    }
}
