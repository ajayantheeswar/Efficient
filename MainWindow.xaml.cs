using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO.Enumeration;
using System.Linq;
using System.Linq.Expressions;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace Win
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.products = new List<ProductRow>();

            backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += backgroundWorker_DoWork;
            //For the display of operation progress to UI.    
            backgroundWorker.ProgressChanged += backgroundWorker_ProgressChanged;
            //After the completation of operation.    
            backgroundWorker.RunWorkerCompleted += backgroundWorker_RunWorkerCompleted;

            backgroundWorker.WorkerReportsProgress = true;

            ProgressBarSticker.Maximum = 100;
            ProgressBarSticker.Value = 0;

            BackDrop.Visibility = Visibility.Collapsed;
            
        }

        String FilePath = "";
        List<ProductRow> products;

        public BackgroundWorker backgroundWorker;

        public class GenrateStickersArguement
        {
            public WorkBookNew wb;
            public string FilePath;
        }

        void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            GenrateStickersArguement Args = e.Argument as GenrateStickersArguement;
            if(Args != null)
            {
                Args.wb.WriteStickerToExcelPrint(Args.FilePath, backgroundWorker);
            }

            
        }

        void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int progress = e.ProgressPercentage;
            int numbersCompleted = (int) (( (decimal) progress / 100) * this.products.Count);
            ProgressMessage.Content = "" + numbersCompleted + " / " + this.products.Count + " Completed" ;
            ProgressBarSticker.Value = progress;
            
        }

        void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            if (e.Cancelled)
            {
                Console.WriteLine("Operation Cancelled");
            }
            else if (e.Error != null)
            {
                Console.WriteLine("Error in Process :" + e.Error);
            }
            else
            {
                MessageBox.Show("Excel file created.");
                BackDrop.Visibility = Visibility.Collapsed;
                
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
       

            OpenFileDialog openFileDialog = new OpenFileDialog();
            
            
            openFileDialog.Filter = "Excel |*.xlsx";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;


            if (openFileDialog.ShowDialog() == true)
            {
                //Get the path of specified file
                FilePath = openFileDialog.FileName;
                string[] FileSection = FilePath.Split('\\');
                FileName.Content = FileSection.Last();
            }
            
        }

        //private void Button_Click_1(object sender, RoutedEventArgs e)
        //{


        //    OpenFileDialog openFileDialog = new OpenFileDialog();


        //    openFileDialog.Filter = "Excel |*.xlsx";
        //    openFileDialog.FilterIndex = 2;
        //    openFileDialog.RestoreDirectory = true;


        //    if (openFileDialog.ShowDialog() == true)
        //    {
        //        //Get the path of specified file
        //        FilePath = openFileDialog.FileName;
        //        string[] FileSection = FilePath.Split('\\');
        //        FileName.Content = FileSection.Last();
        //        List<ProductRow> products = extractData();
        //        RowsCount.Content = products.Count;
        //        WorkBookNew wb = new WorkBookNew(products);

        //        SaveFileDialog saveFileDialog1 = new SaveFileDialog();
        //        saveFileDialog1.Filter = "Excel |*.xlsx";
        //        saveFileDialog1.RestoreDirectory = true;

        //        if (saveFileDialog1.ShowDialog() == true)
        //        {
        //            wb.WriteStickerToExcelPrint(saveFileDialog1.FileName);
        //        }

        //    }

        //}

        public List<ProductRow> extractData()
        {
            List<ProductRow> productList = new List<ProductRow>();
            try
            {
                if(!string.IsNullOrEmpty(FilePath))
                {
                    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                    
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return productList;
                    }

                    Microsoft.Office.Interop.Excel.Workbook xlWorkBookSource;
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetSource;
                    xlWorkBookSource = xlApp.Workbooks.Open(FilePath);
                    // Selected First Worksheet
                    xlWorkSheetSource = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBookSource.Worksheets.get_Item(1);

                    
                    int rowNumber = 1;
                    while(true)
                    {
                        if (rowNumber > 1)
                        {
                            ProductRow product = new ProductRow();
                            product.ProductionNumber = xlWorkSheetSource.Cells[rowNumber, 1].Value.ToString();
                            product.CATNO = xlWorkSheetSource.Cells[rowNumber, 3].Value.ToString();
                            product.QTY = xlWorkSheetSource.Cells[rowNumber, 8].Value.ToString();


                            if (
                                string.IsNullOrEmpty(product.ProductionNumber) ||
                                string.IsNullOrEmpty(product.CATNO) ||
                                string.IsNullOrEmpty(product.QTY)
                                )
                            {
                                break;
                            } else
                            {
                                product.generateQR();
                                productList.Add(product);
                            }
                        }
                        
                        rowNumber++;
                    }

                    xlWorkBookSource.Close(false);
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorkSheetSource);
                    Marshal.ReleaseComObject(xlWorkBookSource);
                    Marshal.ReleaseComObject(xlApp);
                }
            }
            catch(Exception e) {
                // Known Exception - Alternative Required
              
            }
            return productList;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            System.Windows.Window ScanQr = new ScanQR();
            ScanQr.Show();
            this.Close();
        }

        private void Load_Button_Click(object sender, RoutedEventArgs e)
        {
            if(!string.IsNullOrEmpty(FilePath))
            {
                this.products = extractData();
                RowsCount.Content = "Sticker Count : " + products.Count;
            }
            else
            {
                MessageBox.Show("Please Select A File");
            }
            
        }

        private void Generate_Button_Click(object sender, RoutedEventArgs e)
        {
            if(this.products.Count > 0) {
                WorkBookNew wb = new WorkBookNew(products);
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel |*.xlsx";
                saveFileDialog1.RestoreDirectory = true;

                GenrateStickersArguement genrateStickersArguement = new GenrateStickersArguement();
                genrateStickersArguement.wb = wb;
                

                if (saveFileDialog1.ShowDialog() == true)
                {
                    genrateStickersArguement.FilePath = saveFileDialog1.FileName;
                    ProgressBarSticker.Value = 0;
                    BackDrop.Visibility = Visibility.Visible;
                    backgroundWorker.RunWorkerAsync(argument: genrateStickersArguement);
                    
                }
            }
            else
            {
                MessageBox.Show("Please Select A File with atleast 1 Row");
            }
        }
    }
}
