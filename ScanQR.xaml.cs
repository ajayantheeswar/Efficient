using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace Win
{
    /// <summary>
    /// Interaction logic for ScanQR.xaml
    /// </summary>
    public partial class ScanQR : Window
    {

        

        public ScanQR()
        {
            InitializeComponent();
            DataRow = new ObservableCollection<EditableDataViewExcel>();
            ProductData.ItemsSource = DataRow;

            //SearchBox.LostKeyboardFocus += SearchBox_LostKeyboardFocus;
            checkComboBox.ItemsSource = ComboData.ComboBoxData;
            checkComboBox.SelectedIndex = 0;

            

            ProgressBarSticker.Maximum = 100;
            ProgressBarSticker.Value = 0;

            BackDrop.Visibility = Visibility.Collapsed;

            LoadingExcelWorker = new BackgroundWorker();
            LoadingExcelWorker.DoWork += LoadingExcelWorker_DoWork;
            LoadingExcelWorker.ProgressChanged += LoadingExcelWorker_ProgressChanged;
            LoadingExcelWorker.RunWorkerCompleted += LoadingExcelWorker_RunWorkerCompleted;
            LoadingExcelWorker.WorkerReportsProgress = true;

            


            SavingFileWorker = new BackgroundWorker();
            SavingFileWorker.DoWork += SavingFileWorker_DoWork;
            SavingFileWorker.ProgressChanged += SavingFileWorker_ProgressChanged;
            SavingFileWorker.RunWorkerCompleted += SavingFileWorker_RunWorkerCompleted;
            SavingFileWorker.WorkerReportsProgress = true;
        }

        public class LoadWorkerArgument
        {
            public string FilePath;
        }

        public class SavingExcelArgument
        {
            public string FilePath;
            public List<EditableDataViewExcel> dataList;
        }

        public class LoadWorkerResult
        {
            public List<EditableDataViewExcel> DataRowList;

            public LoadWorkerResult(List<EditableDataViewExcel> dataRowList)
            {
                DataRowList = dataRowList;
            }
        }

        public BackgroundWorker LoadingExcelWorker;
        public BackgroundWorker SavingFileWorker;
        

        void LoadingExcelWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            LoadWorkerArgument Args = e.Argument as LoadWorkerArgument;
            
            if (Args != null)
            {
                List<EditableDataViewExcel> DataRowList = EditableDataViewExcel.LoadDataFromExcelSheet(Args.FilePath, LoadingExcelWorker);
                e.Result = DataRowList;
            }


        }
        
        void LoadingExcelWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int progress = e.ProgressPercentage;
           
                
                ProgressMessage.Content = "" + progress + " % " +" Completed";
                ProgressBarSticker.Value = progress;
            
            

        }

        void LoadingExcelWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
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
                List<EditableDataViewExcel> DataRowList = e.Result as List<EditableDataViewExcel>;
                
                this.DataRow.Clear();
                DataRowList.ForEach(row =>
                {
                    this.DataRow.Add(row);
                });
                this.BackDrop.Visibility = Visibility.Collapsed;
                if (DataRowList.Count > 0)
                {
                    this.ProductData.Visibility = Visibility.Visible;
                }
                else
                {
                    this.ProductData.Visibility = Visibility.Collapsed;
                    MessageBox.Show("Rows are Empty. Please Select Proper File.");

                }
                
            }
            ProgressMessage.Content = "";
        }


        void SavingFileWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            SavingExcelArgument Args = e.Argument as SavingExcelArgument;
            if (Args != null)
            {
                EditableDataViewExcel.WriteDataToExcel(this.DataRow, Args.FilePath, this.SavingFileWorker);
            }


        }

        void SavingFileWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int progress = e.ProgressPercentage;
            int numbersCompleted = (int)(((decimal)progress / 100) * this.DataRow.Count);
            ProgressMessage.Content = "" + numbersCompleted + " / " + this.DataRow.Count + " Completed";
            ProgressBarSticker.Value = progress;
        }

        void SavingFileWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
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
                
                this.BackDrop.Visibility = Visibility.Collapsed;
                MessageBox.Show("Excel File Creadted.");

            }
            ProgressMessage.Content = "";
        }



        private void SearchBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if(string.IsNullOrEmpty(SearchBox.Text))
            {
                return;
            }            
            bool isSuccess = this.SearchAndUpdateScanned(SearchBox.Text);
            if (isSuccess)
            {
                this.SearchBox.Text = string.Empty;
                this.SearchBox.Focus();
            }
        }

        public ObservableCollection<EditableDataViewExcel> DataRow
        {
            get;
            set;
        }

        public string FilePath;
        public string SearchString;

        


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
                this.FileName.Content = FileSection.Last();
            }
        }

        // Load Excel
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.FilePath))
            {
                LoadWorkerArgument argument = new LoadWorkerArgument();
                argument.FilePath = this.FilePath;
                BackDrop.Visibility = Visibility.Visible;
                ProgressBarSticker.Value = 0;
                //ProgressLoader.Visibility = Visibility.Collapsed;

                LoadingExcelWorker.RunWorkerAsync(argument: argument);
                //List<EditableDataViewExcel> DataRowList =  EditableDataViewExcel.LoadDataFromExcelSheet(this.FilePath);
                //this.DataRowList = DataRowList;
                //this.DataRow.Clear();

                


                //DataRowList.ForEach(row =>
                //{
                //    this.DataRow.Add(row);
                //});
                ////this.DataRow.
                //if(DataRowList.Count > 0)
                //{
                //    this.ProductData.Visibility = Visibility.Visible;
                //}else
                //{
                //    this.ProductData.Visibility = Visibility.Collapsed;
                //}
            }
            else
            {
                //MessageBox.Show("File Selected isn not empty");
                MessageBox.Show("File is Not Selected");
            }
        }

        //Save Excel
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel |*.xlsx";
            saveFileDialog1.RestoreDirectory = true;

            if(this.ProductData.Items.Count > 0 )
            {
                if (saveFileDialog1.ShowDialog() == true)
                {
                    SavingExcelArgument argument = new SavingExcelArgument();
                    argument.FilePath = saveFileDialog1.FileName;
                    argument.dataList = this.DataRow.ToList();
                    ProgressBarSticker.Value = 0;
                    this.BackDrop.Visibility = Visibility.Visible;
                    SavingFileWorker.RunWorkerAsync(argument: argument);

                    
                }
            }else {

                MessageBox.Show("Please Select File With Valid Data");

            }

            
        }

        public bool SearchAndUpdateScanned(string QRValue)
        {
            string[] qrsplit = QRValue.Trim().Split(' ');
            bool isSearchFound = false;
            if (qrsplit.Length >= 3) {
                string SrNo = qrsplit[0].Trim();
                string CATNO = qrsplit[1].Trim();
                bool hasActualQuantity = qrsplit.Length == 4;
                string ActualQTY = hasActualQuantity ? qrsplit[3].Trim() : "";
                //string SrNo = qrsplit[0];

                int index = 0;
                foreach (var row in this.DataRow)
                {
                    

                    if(row.CATNO.ToLower().Equals(CATNO.ToLower()) &&
                        row.SrNO.ToLower().Equals(SrNo.ToLower()))
                    {
                        ComboData? SelectedItem = checkComboBox.SelectedItem as ComboData;
                        string shortage = string.Empty;
                        isSearchFound = true;

                        row.RowColor = EditableDataViewExcel.Green;

                        if (SelectedItem == null)
                        {
                            return false;
                        }
                        if (SelectedItem.ID == CheckList.Check1)
                        {
                            row.CHECKREMARK1 = DateTime.Now.ToString();
                            
                            if(hasActualQuantity)
                            {
                                row.CHECKACT1 = ActualQTY.ToString();
                                row.CHECKSHORTAGE1 = GetTotalShortage(row.Qty, ActualQTY);
                            } else
                            {
                                row.CHECKACT1 = row.Qty;
                            }
                        }
                        else if (SelectedItem.ID == CheckList.Check2)
                        {
                            row.CHECKREMARK2 = DateTime.Now.ToString();

                            if (hasActualQuantity)
                            {
                                row.CHECKACT2 = ActualQTY.ToString();
                                row.CHECKSHORTAGE2 = GetTotalShortage(row.Qty, ActualQTY);

                            }else
                            {
                                row.CHECKACT2 = row.Qty;
                            }
                        }
                        else if (SelectedItem.ID == CheckList.Check3)
                        {
                            row.CHECKREMARK3 = DateTime.Now.ToString();

                            if (hasActualQuantity)
                            {
                                row.CHECKACT3 = ActualQTY.ToString();
                                row.CHECKSHORTAGE3 = GetTotalShortage(row.Qty, ActualQTY);
                            } else
                            {
                                row.CHECKACT3 = row.Qty;
                            }
                        }
                       this.ProductData.ScrollIntoView(this.ProductData.Items[index]);
                        this.ProductData.UpdateLayout();
                    }
                    else
                    {
                        row.RowColor = EditableDataViewExcel.White;
                    }

                    index++;
                }

                this.ProductData.Items.Refresh();

                if(!isSearchFound)
                {
                    MessageBox.Show("Unable to Find the QR Value , Please Check");
                    return false;
                }
                else
                {
                    return true;
                }
                
            }
            
            

            return false;
        }

        public string GetTotalShortage(string QTY , string ActualQty)
        {
            string Shortage = string.Empty;

            if(string.IsNullOrEmpty(QTY) || string.IsNullOrEmpty(ActualQty))
            {
                return Shortage;
            }

            try
            {
                int qty = 0;
                int.TryParse(QTY, out qty);

                int actualQty = 0;
                int.TryParse(ActualQty, out actualQty);

                Shortage = "" + (qty - actualQty);

                if((qty - actualQty) == 0) { return string.Empty; }
            }
            catch(Exception ex)
            {

            }
            return Shortage;
        }


        private void Mark_Click(object sender, RoutedEventArgs e)
        {
            
            if (string.IsNullOrEmpty(SearchBox.Text.Trim()))
            {
                return;
            }
            bool isSuccess = this.SearchAndUpdateScanned(SearchBox.Text);
            if (isSuccess)
            {
                this.SearchBox.Text = string.Empty;
            }
        }

        private void Capture_Enter_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key.Equals(Key.Enter)) {
                if (string.IsNullOrEmpty(SearchBox.Text))
                {
                    return;
                }
                bool isSuccess = this.SearchAndUpdateScanned(SearchBox.Text);
                if(isSuccess)
                {
                    this.SearchBox.Text = string.Empty;
                }
            }
        }

        public void ProductDataCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            ComboData? SelectedItem = checkComboBox.SelectedItem as ComboData;

            if (e.EditAction == DataGridEditAction.Commit)
            {
                var column = e.Column as DataGridBoundColumn;
                int rowIndex = e.Row.GetIndex();

                TextBox? el = e.EditingElement as TextBox;
                int AcutalNumber = 0;
                bool isValid = int.TryParse(el.Text, out AcutalNumber);
                int QTY = 0;
                
                if (column != null)
                {
                   
                    var bindingPath = (column.Binding as Binding).Path.Path;

                    EditableDataViewExcel DataRow = this.DataRow[rowIndex];
                    int.TryParse(DataRow.Qty, out QTY);

                    int difference = QTY - AcutalNumber;

                    if (bindingPath.Equals("CHECKACT1") && SelectedItem.ID == CheckList.Check1)
                    {
                        
                        if (isValid)
                        {
                            DataRow.CHECKSHORTAGE1 = difference != 0 ? (difference).ToString(): "";
                            DataRow.CHECKREMARK1 = DateTime.Now.ToString();
                        }
                        
                    }else if (bindingPath.Equals("CHECKACT2") && SelectedItem.ID == CheckList.Check2) {
                        if (isValid)
                        {
                            DataRow.CHECKSHORTAGE2 = difference != 0 ? (difference).ToString() : "";
                            DataRow.CHECKREMARK2 = DateTime.Now.ToString();
                        }


                    } else if (bindingPath.Equals("CHECKACT2") && SelectedItem.ID == CheckList.Check3)
                    {
                        if (isValid)
                        {
                            DataRow.CHECKSHORTAGE3 = difference != 0 ? (difference).ToString() : "";
                            DataRow.CHECKREMARK3 = DateTime.Now.ToString();
                        }
                    }

                    
                }
            }
        }
       
    }
}
