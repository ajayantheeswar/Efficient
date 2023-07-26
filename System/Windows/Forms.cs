using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using QRCoder;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Shapes;
using Range = Microsoft.Office.Interop.Excel.Range;
using XlHAlign = Microsoft.Office.Interop.Excel.XlHAlign;

namespace Win
{
    public sealed class TemporaryFile : IDisposable
    {
        public TemporaryFile(string extension)
        {
            Create(System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName().Replace('.', '_') + extension));
        }



        ~TemporaryFile()
        {
            Delete();
        }

        public void Dispose()
        {
            Delete();
            GC.SuppressFinalize(this);
        }

        public string FilePath { get; private set; }

        private void Create(string path)
        {
            FilePath = path;
            using (System.IO.File.Create(FilePath)) { };
        }

        private void Delete()
        {
            if (FilePath == null) return;
            System.IO.File.Delete(FilePath);
            FilePath = null;
        }
    }

    public class ProductRow
    {
        public string ProductionNumber;
        public string CATNO;
        public string QTY;
        public string AbsoluteRowNumber;

        const string ColumnSticker = "M";

        public Bitmap QR;

        public Bitmap StickerImage;

        string Env;

        string projectDirectory;

        public ProductRow()
        {

        }

        public string getQRText()
        {
            return ProductionNumber + " "
                + CATNO + " " + QTY + " ";
        }

        public string getStringSticker()
        {
            return "Prod Number : " + ProductionNumber + "\n"
                + "CATNO : " + CATNO + "\n"
                + "QTY : " + QTY;
        }

        public void CreateSticker()
        {
            Bitmap sticket = new Bitmap(250, 250);
            using (Graphics g = Graphics.FromImage(sticket))
            {
                g.FillRectangle(new SolidBrush(Color.White), 0, 0, 250, 250);
                g.DrawImage(this.QR, 40, 5, 170, 170);
                System.Drawing.Font drawFont = new System.Drawing.Font("Arial", 10);
                SolidBrush drawBrush = new SolidBrush(Color.Black);
                g.DrawString(this.getStringSticker(), drawFont, drawBrush, 10, 175);
            }
            this.StickerImage = sticket;
            //sticket.Save("C:\\Users\\AJAY\\Desktop\\Mama\\" + ProductionNumber + ".png", ImageFormat.Png);

        }

        public void generateQR()
        {
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(this.getQRText(), QRCodeGenerator.ECCLevel.Q);
            QRCode qrCode = new QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(15);
            this.QR = qrCodeImage;
            this.CreateSticker();

        }

        /*
        public void InsertSticker(Microsoft.Office.Interop.Excel.Worksheet activeWorksheet, string rowNumber)
        {
            Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.get_Range(ColumnSticker + rowNumber);

            float Left = (float)((double)oRange.Left);
            float Top = (float)((double)oRange.Top);
            const float ImageSize = 250;

            oRange.RowHeight = ImageSize + 5;
            oRange.ColumnWidth = 50;

            //Range PngOutputColumn = activeWorksheet.get_Range("M" + selected.Rows.Count);
            oRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            oRange.Borders.Color = ColorTranslator.ToOle(Color.Black);

            using (var tempFile = new TemporaryFile(".png"))
            {
                this.StickerImage.Save(tempFile.FilePath);
                Shape image = activeWorksheet.Shapes.AddPicture(tempFile.FilePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left + 2, Top + 2, ImageSize, ImageSize);
                image.Placement = XlPlacement.xlMoveAndSize;
                tempFile.Dispose();
            }


        }
        */
    }

    public class WorkBookNew
    {
        public List<ProductRow> products;

        public WorkBookNew()
        {
            this.products = new List<ProductRow>();
        }

        public WorkBookNew(List<ProductRow> products)
        {
            this.products = products;
        }

        public string previewExcel()
        {
            return null;
        }

        public void WriteStickerToExcel(string FilePath)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = null;

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Open( System.IO.Path.GetDirectoryName(Environment.ProcessPath) + "\\Assets\\template.xlsx");
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "ProductNumber";
            xlWorkSheet.Cells[1, 2] = "CATNO";
            xlWorkSheet.Cells[1, 3] = "QTY";
            xlWorkSheet.Cells[1, 4] = "QR";

            for(int rowNumber = 2 , i = 0 ; i < this.products.Count; i ++ , rowNumber ++)
            {

                xlWorkSheet.Cells[rowNumber, 1] = this.products[i].ProductionNumber;
                xlWorkSheet.Cells[rowNumber, 2] = this.products[i].CATNO;
                xlWorkSheet.Cells[rowNumber, 3] = this.products[i].QTY;
           
                Range oRange = (Range)xlWorkSheet.get_Range("D" + rowNumber);

                float Left = (float)(double)oRange.Left;
                float Top = (float)((double)oRange.Top);
                const float ImageSize = 250;

                oRange.RowHeight = ImageSize + 5;
                oRange.ColumnWidth = 50;

                //Range PngOutputColumn = activeWorksheet.get_Range("M" + selected.Rows.Count);
                oRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                oRange.Borders.Color = ColorTranslator.ToOle(Color.Black);

               

                using (var tempFile = new TemporaryFile(".png"))
                {
                    this.products[i].StickerImage.Save(tempFile.FilePath);
                    Microsoft.Office.Interop.Excel.Shape image = xlWorkSheet.Shapes.AddPicture(tempFile.FilePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left + 2, Top + 2, ImageSize, ImageSize);
                    //AddPicture(tempFile.FilePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left+2, Top+2, ImageSize, ImageSize);
                    image.Placement = XlPlacement.xlMoveAndSize;
                    tempFile.Dispose();
                }
            }

            

            //Here saving the file in xlsx
            xlWorkBook.SaveAs(FilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
            misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel file created");
        }

        public string GetColumnName(int col)
        {
            switch(col)
            {
                case 1: return "A";
                case 2: return "B";

                case 4: return "D";
                case 5: return "E";

                case 7: return "G";
                case 8: return "H";

                case 10: return "J";
                case 11: return "K";
                
                case 13: return "M";
                case 14: return "N";
            }
            return "";
        }

        public void WriteStickerToExcelPrint(string FilePath, BackgroundWorker backgroundWorker)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;

            try
            {

                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Open(System.IO.Path.GetDirectoryName(Environment.ProcessPath) + "\\Assets\\template-beta.xlsx");
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                int i = 0;
                int rowNumber = 2;

                while (i < this.products.Count)
                {
                    int row = rowNumber;
                    int col = 1;


                    if (i % 5 == 0)
                    {
                        col = 1;
                    }
                    else if (i % 5 == 1)
                    {
                        col = 4;
                    }
                    else if (i % 5 == 2)
                    {
                        col = 7;
                    }
                    else if (i % 5 == 3)
                    {
                        col = 10;
                    }
                    else if (i % 5 == 4)
                    {
                        col = 13;
                    }



                    xlWorkSheet.Cells[rowNumber, col + 1] = this.products[i].ProductionNumber;
                    xlWorkSheet.Cells[rowNumber + 1, col + 1] = this.products[i].CATNO;
                    xlWorkSheet.Cells[rowNumber + 2, col + 1] = "QTY: " + this.products[i].QTY;



                    Range oRange1 = (Range)xlWorkSheet.get_Range(GetColumnName(col + 1) + (rowNumber));
                    Range oRange2 = (Range)xlWorkSheet.get_Range(GetColumnName(col + 1) + (rowNumber + 1));
                    Range oRange3 = (Range)xlWorkSheet.get_Range(GetColumnName(col + 1) + (rowNumber + 2));

                    oRange1.Font.Size = 6.5;
                    oRange1.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    oRange2.Font.Size = 6.5;
                    oRange2.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    oRange3.Font.Size = 6.5;
                    oRange3.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    ///Merge 

                    Range oRange = (Range)xlWorkSheet.get_Range(GetColumnName(col) + rowNumber, GetColumnName(col) + (rowNumber + 2));

                    oRange.Merge();

                    float Left = (float)(double)oRange.Left;
                    float Top = (float)((double)oRange.Top);
                    const float ImageSize = 37;

                    //oRange.RowHeight = ImageSize + 3;

                    //Range PngOutputColumn = activeWorksheet.get_Range("M" + selected.Rows.Count);
                    //oRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                    //oRange.Borders.Color = ColorTranslator.ToOle(Color.Black);



                    using (var tempFile = new TemporaryFile(".png"))
                    {
                        this.products[i].QR.Save(tempFile.FilePath);
                        Microsoft.Office.Interop.Excel.Shape image = xlWorkSheet.Shapes.AddPicture(tempFile.FilePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left + 3, Top, ImageSize, ImageSize);
                        //AddPicture(tempFile.FilePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left+2, Top+2, ImageSize, ImageSize);
                        image.Placement = XlPlacement.xlMoveAndSize;
                        tempFile.Dispose();
                    }


                    if ((rowNumber + 3) % 65 == 0)
                    {
                        Range PB = (Range)xlWorkSheet.get_Range("A" + (rowNumber+4));

                        xlWorkSheet.HPageBreaks.Add(PB);
                    }

                    if ((i + 1) % 5 == 0)
                    {
                        rowNumber = rowNumber + 5;
                    }

                    decimal progress = ((i + 1) / (decimal)this.products.Count) * 100;


                    backgroundWorker.ReportProgress((int)progress);
                    i++;
                }

                
                

                //Here saving the file in xlsx
                xlWorkBook.SaveAs(FilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
                misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


                xlWorkBook.Close(false, misValue, misValue);
            }
            catch(Exception Ex)
            {

            }
      
            

            
            Marshal.ReleaseComObject(xlApp);

            
        }
    }
}