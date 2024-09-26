
using System.Drawing;
// using System.Drawing.Imaging;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Tesseract;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;




namespace emails_worker_service.Pdf
{
    public class PdfReaderComponent
    {
        private readonly string _tesseractDataPath;
        public PdfReaderComponent()
        {
            _tesseractDataPath = @"./tessdata/tessdata-main";
        }

        public List<string> ReadPdfAndExtractText(string filePath)
        {
            var extractedTextList = new List<string>();
            if (Path.GetExtension(filePath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    extractedTextList.AddRange(ExtractTextFromDocx(filePath));
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error processing DOCX: " + ex.Message);
                }
            }
            else if (Path.GetExtension(filePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    using (PdfReader reader = new PdfReader(filePath))
                    {
                        using (PdfDocument pdfDoc = new PdfDocument(reader))
                        {
                            for (int pageNumber = 1; pageNumber <= pdfDoc.GetNumberOfPages(); pageNumber++)
                            {
                                var text = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(pageNumber), new SimpleTextExtractionStrategy());
                                if (string.IsNullOrWhiteSpace(text))
                                {
                                    var imagePath = SaveImageFromPage(pdfDoc.GetPage(pageNumber), pageNumber);
                                    text = ExtractTextFromImage(imagePath);
                                    File.Delete(imagePath); // Clean up the image file
                                }
                                extractedTextList.Add(text);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error processing PDF: " + ex.Message);
                }
            }
            else
            {
                Console.WriteLine("Unsupported file format.");
            }

            return extractedTextList;
        }

        private List<string> ExtractTextFromDocx(string filePath)
        {
            List<string> textList = new List<string>();
            try { 
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
                    {
                        Body body = wordDoc.MainDocumentPart.Document.Body;
                        foreach (Paragraph paragraph in body.Elements<Paragraph>())
                        {
                            foreach (Run run in paragraph.Elements<Run>())
                            {
                                foreach (Text text in run.Elements<Text>())
                                {
                                    textList.Add(text.Text);
                                }
                            }
                        }
                    }
                }
                catch (FileNotFoundException)
                {
                    Console.WriteLine("The specified file was not found.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("An error occurred: " + ex.Message);
                }
            
            return textList;
        }

        private string SaveImageFromPage(PdfPage page, int pageNumber)
        {
            var pageDict = page.GetPdfObject();
            var resources = pageDict.GetAsDictionary(PdfName.Resources);
            var xObject = resources.GetAsDictionary(PdfName.XObject);
            var imagePath = string.Empty;

            foreach (var key in xObject.KeySet())
            {
                var pdfObject = xObject.GetAsStream(key);
                if (pdfObject != null && pdfObject.IsStream())
                {
                    var stream = (PdfStream)pdfObject;
                    var bytes = stream.GetBytes();

                    using (var ms = new MemoryStream(bytes))
                    {
                        using(var bitmap = new Bitmap(ms))
                        {
                                imagePath = Path.Combine(Path.GetTempPath(), $"page_{pageNumber}.png");
                                bitmap.Save(imagePath, System.Drawing.Imaging.ImageFormat.Png);
                        }
                    }
                }
            }

            return imagePath;
        }

        private string ExtractTextFromImage(string imagePath)
        {
            try
            {
                using (var engine = new TesseractEngine(_tesseractDataPath, "eng+heb", EngineMode.Default))
                {
                    using (var img = Pix.LoadFromFile(imagePath))
                    {
                        using (var page = engine.Process(img))
                        {
                            return page.GetText();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error extracting text from image: " + ex.Message);
                return string.Empty;
            }
        }

        public Bitmap Resize(Bitmap bmp, int newWidth, int newHeight)
        {

            Bitmap temp = (Bitmap)bmp;

            Bitmap bmap = new Bitmap(newWidth, newHeight, temp.PixelFormat);

            double nWidthFactor = (double)temp.Width / (double)newWidth;
            double nHeightFactor = (double)temp.Height / (double)newHeight;

            double fx, fy, nx, ny;
            int cx, cy, fr_x, fr_y;
            System.Drawing.Color color1 = new System.Drawing.Color();
            System.Drawing.Color color2 = new System.Drawing.Color();
            System.Drawing.Color color3 = new System.Drawing.Color();
            System.Drawing.Color color4 = new System.Drawing.Color();
            byte nRed, nGreen, nBlue;

            byte bp1, bp2;

            for (int x = 0; x < bmap.Width; ++x)
            {
                for (int y = 0; y < bmap.Height; ++y)
                {

                    fr_x = (int)Math.Floor(x * nWidthFactor);
                    fr_y = (int)Math.Floor(y * nHeightFactor);
                    cx = fr_x + 1;
                    if (cx >= temp.Width) cx = fr_x;
                    cy = fr_y + 1;
                    if (cy >= temp.Height) cy = fr_y;
                    fx = x * nWidthFactor - fr_x;
                    fy = y * nHeightFactor - fr_y;
                    nx = 1.0 - fx;
                    ny = 1.0 - fy;

                    color1 = temp.GetPixel(fr_x, fr_y);
                    color2 = temp.GetPixel(cx, fr_y);
                    color3 = temp.GetPixel(fr_x, cy);
                    color4 = temp.GetPixel(cx, cy);

                    // Blue
                    bp1 = (byte)(nx * color1.B + fx * color2.B);

                    bp2 = (byte)(nx * color3.B + fx * color4.B);

                    nBlue = (byte)(ny * (double)(bp1) + fy * (double)(bp2));

                    // Green
                    bp1 = (byte)(nx * color1.G + fx * color2.G);

                    bp2 = (byte)(nx * color3.G + fx * color4.G);

                    nGreen = (byte)(ny * (double)(bp1) + fy * (double)(bp2));

                    // Red
                    bp1 = (byte)(nx * color1.R + fx * color2.R);

                    bp2 = (byte)(nx * color3.R + fx * color4.R);

                    nRed = (byte)(ny * (double)(bp1) + fy * (double)(bp2));

                    bmap.SetPixel(x, y, System.Drawing.Color.FromArgb
            (255, nRed, nGreen, nBlue));
                }
            }



            bmap = SetGrayscale(bmap);
            bmap = RemoveNoise(bmap);

            return bmap;

        }

        public Bitmap SetGrayscale(Bitmap img)
        {

            Bitmap temp = (Bitmap)img;
            Bitmap bmap = (Bitmap)temp.Clone();
            System.Drawing.Color c;
            for (int i = 0; i < bmap.Width; i++)
            {
                for (int j = 0; j < bmap.Height; j++)
                {
                    c = bmap.GetPixel(i, j);
                    byte gray = (byte)(.299 * c.R + .587 * c.G + .114 * c.B);

                    bmap.SetPixel(i, j, System.Drawing.Color.FromArgb(gray, gray, gray));
                }
            }
            return (Bitmap)bmap.Clone();

        }

        public Bitmap RemoveNoise(Bitmap bmap)
        {

            for (var x = 0; x < bmap.Width; x++)
            {
                for (var y = 0; y < bmap.Height; y++)
                {
                    var pixel = bmap.GetPixel(x, y);
                    if (pixel.R < 162 && pixel.G < 162 && pixel.B < 162)
                        bmap.SetPixel(x, y, System.Drawing.Color.Black);
                    else if (pixel.R > 162 && pixel.G > 162 && pixel.B > 162)
                        bmap.SetPixel(x, y, System.Drawing.Color.White);
                }
            }

            return bmap;
        }
    }

  /*  class Program
    {
        static void Main(string[] args)
        {
            var tesseractDataPath = @".\tessdata"; // Path to the 'tessdata' folder containing language data files
            var pdfFilePath = @"C:/Users/recruitment/Desktop/sample.pdf";

            var pdfReader = new PdfReaderComponent(tesseractDataPath);
            var extractedTexts = pdfReader.ReadPdfAndExtractText(pdfFilePath);

            foreach (var text in extractedTexts)
            {
                Console.WriteLine("Extracted Text: \n" + text);
            }
        }
    }*/
}
