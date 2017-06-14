using System;
using System.Data.OleDb;
using System.Data;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
//using iTextSharp.xtra.iTextSharp.text.pdf.pdfcleanup;


namespace Data
{
    public class excelclass
    {

        public partial class Footer : PdfPageEventHelper

        {

            public override void OnEndPage(PdfWriter writer, Document doc)

            {

                Paragraph footer = new Paragraph("THANK YOU", FontFactory.GetFont(FontFactory.TIMES, 10, iTextSharp.text.Font.NORMAL));

                footer.Alignment = Element.ALIGN_RIGHT;

                PdfPTable footerTbl = new PdfPTable(1);

                footerTbl.TotalWidth = 300;

                footerTbl.HorizontalAlignment = Element.ALIGN_CENTER;

                PdfPCell cell = new PdfPCell(footer);

                cell.Border = 0;

                cell.PaddingLeft = 10;

                footerTbl.AddCell(cell);

                footerTbl.WriteSelectedRows(0, -1, 415, 30, writer.DirectContent);

            }

        }


        public void AddPageNumber(string str, string displaytext)
        {
            byte[] bytes = File.ReadAllBytes(str);
            
            iTextSharp.text.Font blackFont = FontFactory.GetFont("Arial", 6, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);

            using (var stream = new MemoryStream())
            {
                stream.Flush();
                System.Windows.Forms.Application.DoEvents();
                PdfReader reader = new PdfReader(bytes);
                PdfReader.unethicalreading = true;
                Paragraph p = new Paragraph();
                Document doc = new Document();
                using (PdfStamper stamper = new PdfStamper(reader, stream))
                {
                    System.Windows.Forms.Application.DoEvents();

                    //PdfContentByte canvas = stamper.GetOverContent(1);
                    //iTextSharp.text.Rectangle size = reader.GetPageSizeWithRotation(1);
                    int pages = reader.NumberOfPages;
                    for (int i = 1; i <= pages; i++)
                    {
                        iTextSharp.text.Rectangle mediabox = reader.GetPageSize(i);
                        float x, y;
                        x = mediabox.Height;
                        y = mediabox.Width;
                      
                        iTextSharp.text.Rectangle cropBox = reader.GetCropBox(i);
                        cropBox.GetRectangle(y-1000,x-50);
                        iTextSharp.text.Rectangle bottomRight = new iTextSharp.text.Rectangle(cropBox.GetRight(199), cropBox.Bottom, cropBox.Right, cropBox.GetBottom(18));
                        //TextField tf = new TextField(stamper.Writer, new iTextSharp.text.Rectangle(0, 0, 300, 100), displaytext);
                        ////Change the orientation of the text
                        //tf.Rotation = 0;

                        EmptyTextBoxSimple(stamper, i, bottomRight, BaseColor.WHITE);
                        ColumnText columnText = GenerateTextBox(stamper, i, bottomRight);
                        columnText.AddText(new Phrase(displaytext.ToString(), blackFont));
                        columnText.Go();
                    }
                }
                  
                    bytes = stream.ToArray();
                //stream.Flush();
            }
            System.Threading.Thread.Sleep(200);
            File.WriteAllBytes(str, bytes);
        }
       
        public void WriteToLog(string msg, string stkTrace, string title)
        {
            if (!(System.IO.Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\Errors\\")))
            {
                System.IO.Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\Errors\\");
            }
            FileStream fs = new FileStream(System.Windows.Forms.Application.StartupPath + "\\Errors\\errlog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter s = new StreamWriter(fs);
            s.Close();
            fs.Close();
            FileStream fs1 = new FileStream(System.Windows.Forms.Application.StartupPath + "\\Errors\\errlog.txt", FileMode.Append, FileAccess.Write);
            StreamWriter s1 = new StreamWriter(fs1);
            s1.Write("Title: " + title);
            s1.Write("Message: " + msg);
            s1.Write("StackTrace: " + stkTrace);
            s1.Write("Date/Time: " + DateTime.Now.ToString());
            s1.Write
                ("============================================");
            s1.Close();
            fs1.Close();
        }
        void EmptyTextBoxSimple(PdfStamper stamper, int pageNumber, iTextSharp.text.Rectangle boxArea, BaseColor fillColor)
        {
            PdfContentByte canvas = stamper.GetOverContent(pageNumber);
            canvas.SaveState();
            canvas.SetColorFill(fillColor);
            boxArea.BorderWidth = 2;
            boxArea.BorderColor = new BaseColor(1, 2, 3);
            canvas.Rectangle(boxArea.Left, boxArea.Bottom, boxArea.Width, boxArea.Height-10);
            canvas.Fill();
            canvas.RestoreState();
        }
        ColumnText GenerateTextBox(PdfStamper stamper, int pageNumber, iTextSharp.text.Rectangle boxArea)
        {
            PdfContentByte canvas = stamper.GetOverContent(pageNumber);
            ColumnText columnText = new ColumnText(canvas);
            columnText.SetSimpleColumn(boxArea);
            return columnText;
        }



    }
}

