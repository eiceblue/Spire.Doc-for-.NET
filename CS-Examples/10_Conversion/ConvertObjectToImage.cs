using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;
using Spire.Doc.Fields;
using System.Drawing.Imaging;

namespace ConvertObjectToImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a document
            Document document = new Document();
            //Load file
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ConvertObjectToImage.docx");
            //Get the first section
            Section section = document.Sections[0];
            //Get body of section
            Body body = section.Body;

            //Get the first paragraph
            Paragraph paragraph = body.Paragraphs[0]; 
            Image image = ConvertParagraphToImage(paragraph);
            image.Save("ConvertParagraphToImage.png",ImageFormat.Png);

            //Get the first table
            Table table = body.Tables[0] as Table;
            image = ConvertTableToImage(table);
            image.Save("ConvertTableToImage.jpg", ImageFormat.Jpeg);

            //Get the first row of the first table
            TableRow row = table.Rows[0];
            image = ConvertTableRowToImage(row);
            image.Save("ConvertTableRowToImage.bmp", ImageFormat.Bmp);

            //Get the first cell of the first row
            TableCell cell = row.Cells[0];
            image = ConvertTableCellToImage(cell);
            image.Save("ConvertTableCellToImage.png", ImageFormat.Png);

            //Get a shape
            int i = 0;
            foreach (Paragraph p in section.Paragraphs)
            {
                foreach (DocumentObject obj in p.ChildObjects)
                {
                    if (obj.DocumentObjectType == DocumentObjectType.Shape)
                    {
                        image = ConvertShapeToImage(obj as ShapeObject);
                        image.Save(String.Format("ConvertShapeToImage-{0}.png",i), ImageFormat.Png);
                        i++;
                    }
                }
            }
           
        }
        private Image ConvertParagraphToImage(Paragraph obj)
        {
            Document doc = new Document();
            Section section = doc.AddSection();
     
            section.Body.ChildObjects.Add(obj.Clone());
            Image image = doc.SaveToImages(0, ImageType.Bitmap);
            doc.Close();    
            return CutImageWhitePart(image as Bitmap, 1); 
        }
        private Image ConvertTableToImage(Table obj)
        {
            Document doc = new Document();
            Section section = doc.AddSection();

            section.Body.ChildObjects.Add(obj.Clone());

            Image image = doc.SaveToImages(0, ImageType.Bitmap);
            doc.Close();    
            return CutImageWhitePart(image as Bitmap, 1); 
        }
        private Image ConvertTableRowToImage(TableRow obj)
        {
            Document doc = new Document();
            Section section = doc.AddSection();
            Table table = section.AddTable();
            table.Rows.Add(obj.Clone());
            Image image = doc.SaveToImages(0, ImageType.Bitmap);
            doc.Close();    
            return CutImageWhitePart(image as Bitmap, 1); 
        }
        private Image ConvertTableCellToImage(TableCell obj)
        {
            Document doc = new Document();
            Section section = doc.AddSection();
            Table table = section.AddTable();
            table.AddRow().Cells.Add(obj.Clone());
            Image image = doc.SaveToImages(0, ImageType.Bitmap);
            doc.Close();    
            return CutImageWhitePart(image as Bitmap, 1); 
        }
        private Image ConvertShapeToImage(ShapeObject obj)
        {
            Document doc = new Document();
            Section section = doc.AddSection();
            section.AddParagraph().ChildObjects.Add(obj.Clone());
            MemoryStream ms = new MemoryStream();
            doc.SaveToStream(ms, FileFormat.Docx);
            doc.LoadFromStream(ms,FileFormat.Docx);
            Image image = doc.SaveToImages(0, ImageType.Bitmap);
            ms.Close();
            doc.Close();          
            return CutImageWhitePart(image as Bitmap, 1); 
        }
        public Image CutImageWhitePart(Bitmap bmp, int WhiteBarRate)
        {
            int top = 0, left = 0;
            int right = bmp.Width, bottom = bmp.Height;
            Color white = Color.White;

            for (int i = 0; i < bmp.Height; i++)
            {
                bool find = false;
                for (int j = 0; j < bmp.Width; j++)
                {
                    Color c = bmp.GetPixel(j, i);
                    if (IsWhite(c))
                    {
                        top = i;
                        find = true;
                        break;
                    }
                }
                if (find) break;
            }

            for (int i = 0; i < bmp.Width; i++)
            {
                bool find = false;
                for (int j = top; j < bmp.Height; j++)
                {
                    Color c = bmp.GetPixel(i, j);
                    if (IsWhite(c))
                    {
                        left = i;
                        find = true;
                        break;
                    }
                }
                if (find) break; ;
            }

            for (int i = bmp.Height - 1; i >= 0; i--)
            {
                bool find = false;
                for (int j = left; j < bmp.Width; j++)
                {
                    Color c = bmp.GetPixel(j, i);
                    if (IsWhite(c))
                    {
                        bottom = i;
                        find = true;
                        break;
                    }
                }
                if (find) break;
            }

            for (int i = bmp.Width - 1; i >= 0; i--)
            {
                bool find = false;
                for (int j = 0; j <= bottom; j++)
                {
                    Color c = bmp.GetPixel(i, j);
                    if (IsWhite(c))
                    {
                        right = i;
                        find = true;
                        break;
                    }
                }
                if (find) break;
            }
            int iWidth = right - left;
            int iHeight = bottom - left;
            int blockWidth = Convert.ToInt32(iWidth * WhiteBarRate / 100);
            bmp = Cut(bmp, left - blockWidth, top - blockWidth, right - left + 2 * blockWidth, bottom - top + 2 * blockWidth);

            return bmp;

        }
        public Bitmap Cut(Bitmap b, int StartX, int StartY, int iWidth, int iHeight)
        {
            if (b == null)
            {
                return null;
            }
            int w = b.Width;
            int h = b.Height;
            if (StartX >= w || StartY >= h)
            {
                return null;
            }
            if (StartX + iWidth > w)
            {
                iWidth = w - StartX;
            }
            if (StartY + iHeight > h)
            {
                iHeight = h - StartY;
            }
            try
            {
                Bitmap bmpOut = new Bitmap(iWidth, iHeight, PixelFormat.Format24bppRgb);
                Graphics g = Graphics.FromImage(bmpOut);
                g.DrawImage(b, new Rectangle(0, 0, iWidth, iHeight), new Rectangle(StartX, StartY, iWidth, iHeight), GraphicsUnit.Pixel);
                g.Dispose();
                return bmpOut;
            }
            catch
            {
                return null;
            }
        }


        public bool IsWhite(Color c)
        {
            if (c.R < 245 || c.G < 245 || c.B < 245)
                return true;
            else return false;
        }
    }
}
