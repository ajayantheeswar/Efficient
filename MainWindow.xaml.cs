using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.Metrics;
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

            LoadingExcelWorker = new BackgroundWorker();
            LoadingExcelWorker.DoWork += LoadingExcelWorker_DoWork;
            LoadingExcelWorker.ProgressChanged += LoadingExcelWorker_ProgressChanged;
            LoadingExcelWorker.RunWorkerCompleted += LoadingExcelWorker_RunWorkerCompleted;
            LoadingExcelWorker.WorkerReportsProgress = true;

        }

        public class LoadWorkerArgument
        {
            public string? FilePath;
        }

        public class LoadWorkerResult
        {
            public List<ProductRow> DataRowList;

            public LoadWorkerResult(List<ProductRow> dataRowList)
            {
                DataRowList = dataRowList;
            }
        }

        private void LoadingExcelWorker_RunWorkerCompleted(object? sender, RunWorkerCompletedEventArgs e)
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
                List<ProductRow> DataRowList = e.Result as List<ProductRow>;

                this.products.Clear();
                DataRowList.ForEach(row =>
                {
                    this.products.Add(row);
                });
                RowsCount.Content = "Sticker Count : " + products.Count;
                this.BackDrop.Visibility = Visibility.Collapsed;
                if (!(DataRowList.Count > 0))
                {
                    MessageBox.Show("Rows are Empty. Please Select Proper File.");
                }
                

            }
            ProgressMessage.Content = "";
            ProgressBarSticker.Value = 0;
        }

        private void LoadingExcelWorker_ProgressChanged(object? sender, ProgressChangedEventArgs e)
        {
            int progress = e.ProgressPercentage;
            ProgressMessage.Content = "" + progress + " % " + " Completed";
            ProgressBarSticker.Value = progress;
        }

        private void LoadingExcelWorker_DoWork(object? sender, DoWorkEventArgs e)
        {
            LoadWorkerArgument Args = e.Argument as LoadWorkerArgument;

            if (Args != null)
            {
                List<ProductRow> DataRowList = ExtractData(Args.FilePath,LoadingExcelWorker);
                e.Result = DataRowList;
            }
        }

        String FilePath = "";
        List<ProductRow> products;

        public BackgroundWorker backgroundWorker;
        public BackgroundWorker LoadingExcelWorker;

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

       
        public List<ProductRow> ExtractData(string FilePath, BackgroundWorker LoadingBackGroundWorker)
        {
            List<ProductRow> productList = new List<ProductRow>();

            Microsoft.Office.Interop.Excel.Workbook xlWorkBookSource = null;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetSource = null;
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                if(!string.IsNullOrEmpty(FilePath))
                {
                    

                    
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return productList;
                    }

                    xlWorkBookSource = xlApp.Workbooks.Open(FilePath);
                    // Selected First Worksheet
                    xlWorkSheetSource = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBookSource.Worksheets.get_Item(1);

                    int totalRows = xlWorkSheetSource.UsedRange.Columns[1, Type.Missing].Rows.Count;
                    int count = 0;
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

                                if (totalRows > 0)
                                {
                                    int progress = (int)(((count + 1) / (decimal)totalRows) * 100);
                                    LoadingBackGroundWorker.ReportProgress(progress);
                                }


                                count++;
                            }
                        }
                        
                        rowNumber++;
                    }

                    
                }
            }
            catch(Exception e) {
                // Known Exception - Alternative Required
                if(xlWorkBookSource != null)
                    xlWorkBookSource.Close(false);
                if(xlApp != null)
                    xlApp.Quit();

                if (xlWorkSheetSource != null) Marshal.ReleaseComObject(xlWorkSheetSource);
                if (xlWorkBookSource != null) Marshal.ReleaseComObject(xlWorkBookSource);
                Marshal.ReleaseComObject(xlApp);
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
                LoadWorkerArgument Args = new LoadWorkerArgument();
                Args.FilePath = FilePath;

                this.BackDrop.Visibility = Visibility.Visible;
                LoadingExcelWorker.RunWorkerAsync(argument: Args);


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
