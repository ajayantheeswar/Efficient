using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Win
{
    public class ComboData
    {
       public CheckList ID;
       public string? Text;

    public override string ToString()
        {
            return Text;
        }

        public static ComboData[] ComboBoxData = new ComboData[]
            {
                new ComboData()
                {
                    ID = CheckList.Check1,
                    Text = "CHECK 1"
                },
                new ComboData()
                {
                    ID = CheckList.Check2,
                    Text = "CHECK 2"
                },
                new ComboData()
                {
                    ID = CheckList.Check3,
                    Text = "CHECK 3"
                },
            };

    }


    public enum CheckList
    {
        Check1,
        Check2,
        Check3
    }

    public class EditableDataViewExcel : INotifyPropertyChanged // implements
    {
        public const string Blue = "Blue"; 
        public const string White = "White"; 
        
        public string _ID;
        private string rowColor;
        private string srNO;
        private string raw;
        private string cATNO;
        private string noOfBends;
        private string catDesc;
        private string netWtScrap;
        private string compLoc;
        private string qty;
        
        private string checkact1;
        private string checkremark1;
        private string checkshortage1;

        private string checkact2;
        private string checkremark2;
        private string checkshortage2;

        private string checkact3;
        private string checkremark3;
        private string checkshortage3;

        public static string[] HeaderList = new string[] {
            "SR.NO", "RAW", "CAT NO", "No.of.Bends", "CAT DESC", "NET WT. + SCRAP", "COMP LOG", "QTY", "CHECK 1 ACT", "CHECK 1 REMARK","CHECK 1 SHORTAGE" ,"CHECK 2 ACT", "CHECK 2 REMARK","CHECK 2 SHORTAGE","CHECK 3 ACT", "CHECK 3 REMARK","CHECK 3 SHORTAGE" };

        public event PropertyChangedEventHandler? PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public string RowColor
        {
            get => rowColor; set
            {
                if (!value.Equals(this.RowColor))
                {
                    rowColor = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string SrNO
        {
            get => srNO; set
            {
                if (!value.Equals(this.SrNO))
                {
                    srNO = value;
                    NotifyPropertyChanged();
                }

            }
        }
        public string Raw { get => raw; set
            {
                if (!value.Equals(this.raw))
                {
                    raw = value;
                    NotifyPropertyChanged();
                }

            }
        }
        public string CATNO { get => cATNO; set
            {
                if (!value.Equals(this.cATNO))
                {
                    cATNO = value;
                    NotifyPropertyChanged();
                }

            }
        }
        public string NoOfBends { get => noOfBends; set
            {
                if (!value.Equals(this.noOfBends))
                {
                    noOfBends = value;
                    NotifyPropertyChanged();
                }

            }
        }
        public string CatDesc { get => catDesc; set
            {
                if (!value.Equals(this.catDesc))
                {
                    catDesc = value;
                    NotifyPropertyChanged();
                }

            }
        }
        public string NetWtScrap { get => netWtScrap; set
            {
                if (!value.Equals(this.netWtScrap))
                {
                    netWtScrap = value;
                    NotifyPropertyChanged();
                }

            }
        }
        public string CompLoc { get => compLoc; set
            {
                if (!value.Equals(this.compLoc))
                {
                    compLoc = value;
                    NotifyPropertyChanged();
                }

            }
        }
        public string Qty { get => qty; set
            {
                if (!value.Equals(this.qty))
                {
                    qty = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string CHECKACT1
        {
            get => checkact1; set
            {
                if (!value.Equals(this.checkact1))
                {
                    checkact1 = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string CHECKREMARK1
        {
            get => checkremark1; set
            {
                if (!value.Equals(this.checkremark1))
                {
                    checkremark1 = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string CHECKSHORTAGE1
        {
            get => checkshortage1; set
            {
                if (!value.Equals(this.checkshortage1))
                {
                    checkshortage1 = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string CHECKACT2
        {
            get => checkact2; set
            {
                if (!value.Equals(this.checkact2))
                {
                    checkact2 = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string CHECKREMARK2
        {
            get => checkremark2; set
            {
                if (!value.Equals(this.checkremark2))
                {
                    checkremark2 = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string CHECKSHORTAGE2
        {
            get => checkshortage2; set
            {
                if (!value.Equals(this.checkshortage2))
                {
                    checkshortage2 = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string CHECKACT3
        {
            get => checkact3; set
            {
                if (!value.Equals(this.checkact3))
                {
                    checkact3 = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string CHECKREMARK3
        {
            get => checkremark3; set
            {
                if (!value.Equals(this.checkremark3))
                {
                    checkremark3 = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string CHECKSHORTAGE3
        {
            get => checkshortage3; set
            {
                if (!value.Equals(this.checkshortage3))
                {
                    checkshortage3 = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public const string Green = "Green";

        public static bool WriteDataToExcel(ObservableCollection<EditableDataViewExcel> Data, string FilePath,BackgroundWorker SaveFileWorker)
        {
            bool isSuccess = false;
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return false;
                }


                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add();
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                // Writting File Header
                for(int i =0;i< HeaderList.Length;i++) 
                {
                    xlWorkSheet.Cells[1, i+1] = HeaderList[i];
                    ((Range)xlWorkSheet.Cells[1, i+1]).Font.Bold = true;
                    ((Range)xlWorkSheet.Cells[1, i+1]).Columns.AutoFit();
                }
                
                

                //Writting Data
                int RowNumber = 2;
                int count = 0;
                Data.ToList().ForEach((product) =>
                {
                    xlWorkSheet.Cells[RowNumber, 1] = product.SrNO;
                    xlWorkSheet.Cells[RowNumber, 2] = product.Raw;
                    xlWorkSheet.Cells[RowNumber, 3] = product.CATNO;
                    xlWorkSheet.Cells[RowNumber, 4] = product.NoOfBends;
                    xlWorkSheet.Cells[RowNumber, 5] = product.CatDesc; 
                    xlWorkSheet.Cells[RowNumber, 6] = product.NetWtScrap; 
                    xlWorkSheet.Cells[RowNumber, 7] = product.CompLoc; 
                    xlWorkSheet.Cells[RowNumber, 8] = product.Qty;
                     
                    xlWorkSheet.Cells[RowNumber, 9] = product.CHECKACT1; 
                    xlWorkSheet.Cells[RowNumber, 10] = product.CHECKREMARK1;
                    xlWorkSheet.Cells[RowNumber, 11] = product.CHECKSHORTAGE1;

                    xlWorkSheet.Cells[RowNumber, 12] = product.CHECKACT2;
                    xlWorkSheet.Cells[RowNumber, 13] = product.CHECKREMARK2;
                    xlWorkSheet.Cells[RowNumber, 14] = product.CHECKSHORTAGE2;

                    xlWorkSheet.Cells[RowNumber, 15] = product.CHECKACT3;
                    xlWorkSheet.Cells[RowNumber, 16] = product.CHECKREMARK3;
                    xlWorkSheet.Cells[RowNumber, 17] = product.CHECKSHORTAGE3;

                    
                    decimal progress = ((count + 1) / (decimal)Data.Count) * 100;


                    SaveFileWorker.ReportProgress((int)progress);

                    RowNumber++;
                    count++;
                });

                for (int i = 0; i < HeaderList.Length; i++)
                {
                    ((Range)xlWorkSheet.Cells[1, i + 1]).Font.Bold = true;
                    ((Range)xlWorkSheet.Cells[1, i + 1]).ColumnWidth += 10;
                }

                xlWorkBook.SaveAs(FilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
                misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                
            }
            catch(Exception ex)
            {

            }
            return isSuccess;
        }
       

        public static List<EditableDataViewExcel> LoadDataFromExcelSheet(string FilePath,BackgroundWorker LoadingBackGroundWorker)
        {
            List<EditableDataViewExcel> editableDataViewExcels = new List<EditableDataViewExcel>();

            try
            {
                if (!string.IsNullOrEmpty(FilePath))
                {
                    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();


                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return editableDataViewExcels;
                    }

                    Microsoft.Office.Interop.Excel.Workbook xlWorkBookSource;
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetSource;
                    xlWorkBookSource = xlApp.Workbooks.Open(FilePath);
                    // Selected First Worksheet
                    xlWorkSheetSource = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBookSource.Worksheets.get_Item(1);

                    int totalRows = xlWorkSheetSource.UsedRange.Columns[1, Type.Missing].Rows.Count;

                   

                    int rowNumber = 1;

                    int count = 0;
                    while (true)
                    {
                        if (rowNumber > 1)
                        {
                            EditableDataViewExcel product = new EditableDataViewExcel();
                            product._ID = new Guid().ToString();
                            product.RowColor = White;
                            product.SrNO = getCellsValue(xlWorkSheetSource, rowNumber, 1);
                            product.Raw = getCellsValue(xlWorkSheetSource, rowNumber, 2);
                            product.CATNO = getCellsValue(xlWorkSheetSource, rowNumber, 3);
                            product.NoOfBends = getCellsValue(xlWorkSheetSource, rowNumber, 4);
                            product.CatDesc = getCellsValue(xlWorkSheetSource, rowNumber, 5);
                            product.NetWtScrap = getCellsValue(xlWorkSheetSource, rowNumber, 6);
                            product.CompLoc = getCellsValue(xlWorkSheetSource, rowNumber, 7);
                            product.Qty = getCellsValue(xlWorkSheetSource, rowNumber, 8);

                            product.CHECKACT1 = getCellsValue(xlWorkSheetSource, rowNumber, 9); 
                            product.CHECKREMARK1 = getCellsValue(xlWorkSheetSource, rowNumber, 10); 
                            product.CHECKSHORTAGE1 = getCellsValue(xlWorkSheetSource, rowNumber, 11);

                            product.CHECKACT2 = getCellsValue(xlWorkSheetSource, rowNumber, 12);
                            product.CHECKREMARK2 = getCellsValue(xlWorkSheetSource, rowNumber, 13);
                            product.CHECKSHORTAGE2 = getCellsValue(xlWorkSheetSource, rowNumber, 14);

                            product.CHECKACT3 = getCellsValue(xlWorkSheetSource, rowNumber, 15);
                            product.CHECKREMARK3 = getCellsValue(xlWorkSheetSource, rowNumber, 16);
                            product.CHECKSHORTAGE3 = getCellsValue(xlWorkSheetSource, rowNumber, 17);


                            if (
                                string.IsNullOrEmpty(product.SrNO) ||
                                string.IsNullOrEmpty(product.CATNO) ||
                                string.IsNullOrEmpty(product.Qty)
                                )
                            {
                                break;
                            }
                            else
                            {
                                editableDataViewExcels.Add(product);
                                
                                if(totalRows > 0)
                                {
                                    int progress = (int)(((count+1) / (decimal)totalRows)*100);
                                    LoadingBackGroundWorker.ReportProgress(progress);
                                }
                                
                                
                                count++;
                            }
                        }

                        rowNumber++;
                    }

                    xlWorkBookSource.Close(false);
                    xlApp.Quit();

                   
                    string value = "Rows Loaded.";
                    
                    Marshal.ReleaseComObject(xlWorkSheetSource);
                    Marshal.ReleaseComObject(xlWorkBookSource);
                    Marshal.ReleaseComObject(xlApp);
                }
            }
            catch (Exception e)
            {
                // Known Exception - Alternative Required

            }
            return editableDataViewExcels;
        }

        private static string getCellsValue(Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetSource, int row,int column)
        {
            string value = string.Empty;
            try
            {
                value = xlWorkSheetSource.Cells[row, column].Value.ToString();
            }
            catch (Exception ex)
            {

            }
            return value;
        }
    }
}
