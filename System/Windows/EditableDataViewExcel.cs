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
                    Text = "FAB"
                },
                new ComboData()
                {
                    ID = CheckList.Check2,
                    Text = "POWDER COATING"
                },
                new ComboData()
                {
                    ID = CheckList.Check3,
                    Text = "HANDOVER"
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
        private string brdNo;
        private string raw;
        private string cATNO;
        private string noOfBends;
        private string catDesc;
        private string netWtScrap;
        private string compLoc;
        private string qty;
        
        private string fabact;
        private string fabtime;
        private string fabshrt;
        private string fabchkby;

        private string pcact;
        private string pctime;
        private string pcshrt;
        private string pcchkby;

        private string hdact;
        private string hdtime;
        private string hdshrt;
        private string hdchkby;

        public static string[] HeaderList = new string[] {
            "BRD.NO", "RAW", "CAT NO", "No.of.Bends", "CAT DESC", "NET WT. + SCRAP", "COMP LOG", "QTY",
            "FAB ACT", "FAB TIME","FAB SHORTAGE","FAB CHK BY",
            "PC ACT", "PC TIME","PC SHORTAGE","PC CHK BY" ,
            "H/O ACT", "H/O TIME","H/O SHORTAGE","H/O CHK BY" , };

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

        public string BRDNO
        {
            get => brdNo; set
            {
                if (!value.Equals(this.brdNo))
                {
                    brdNo = value;
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

        public string FABACT
        {
            get => fabact; set
            {
                if (!value.Equals(this.fabact))
                {
                    fabact = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string FABTIME
        {
            get => fabtime; set
            {
                if (!value.Equals(this.fabtime))
                {
                    fabtime = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string FABSHORTAGE
        {
            get => fabshrt; set
            {
                if (!value.Equals(this.fabshrt))
                {
                    fabshrt = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string FABCHKBY
        {
            get => fabchkby; set
            {
                if (!value.Equals(this.fabchkby))
                {
                    fabchkby = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string PCACT
        {
            get => pcact; set
            {
                if (!value.Equals(this.pcact))
                {
                    pcact = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string PCTIME
        {
            get => pctime; set
            {
                if (!value.Equals(this.pctime))
                {
                    pctime = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string PCSHRT
        {
            get => pcshrt; set
            {
                if (!value.Equals(this.pcshrt))
                {
                    pcshrt = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string PCCHKBY
        {
            get => pcchkby; set
            {
                if (!value.Equals(this.pcchkby))
                {
                    pcchkby = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string HDACT
        {
            get => hdact; set
            {
                if (!value.Equals(this.hdact))
                {
                    hdact = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string HDTIME
        {
            get => hdtime; set
            {
                if (!value.Equals(this.hdtime))
                {
                    hdtime = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string HDSHRT
        {
            get => hdshrt; set
            {
                if (!value.Equals(this.hdshrt))
                {
                    hdshrt = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public string HDCHKBY
        {
            get => hdchkby; set
            {
                if (!value.Equals(this.hdchkby))
                {
                    hdchkby = value;
                    NotifyPropertyChanged();
                }

            }
        }

        public const string Yellow = "Yellow";

        public const string Green = "Green";

        public static bool WriteDataToExcel(ObservableCollection<EditableDataViewExcel> Data, string FilePath,BackgroundWorker SaveFileWorker)
        {
            bool isSuccess = false;


            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = null;
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();


            try
            {
                

                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return false;
                }



                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add();
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                // Writting File Header
                for(int i =0;i< HeaderList.Length;i++) 
                {
                    xlWorkSheet.Cells[1, i+1] = HeaderList[i];
                    ((Range)xlWorkSheet.Cells[1, i+1]).Font.Bold = true;
                    ((Range)xlWorkSheet.Cells[1, i + 1]).WrapText = true;
                    ((Range)xlWorkSheet.Cells[1, i + 1]).Columns.Font.Size = 8;
                    //((Range)xlWorkSheet.Cells[1, i+1]).Columns.AutoFit();
                }
                
                

                //Writting Data
                int RowNumber = 2;
                int count = 0;
                Data.ToList().ForEach((product) =>
                {
                    xlWorkSheet.Cells[RowNumber, 1] = product.BRDNO;
                    xlWorkSheet.Cells[RowNumber, 2] = product.Raw;
                    xlWorkSheet.Cells[RowNumber, 3] = product.CATNO;
                    xlWorkSheet.Cells[RowNumber, 4] = product.NoOfBends;
                    xlWorkSheet.Cells[RowNumber, 5] = product.CatDesc; 
                    xlWorkSheet.Cells[RowNumber, 6] = product.NetWtScrap; 
                    xlWorkSheet.Cells[RowNumber, 7] = product.CompLoc; 
                    xlWorkSheet.Cells[RowNumber, 8] = product.Qty;
                     
                    xlWorkSheet.Cells[RowNumber, 9] = product.FABACT; 
                    xlWorkSheet.Cells[RowNumber, 10] = product.FABTIME;
                    xlWorkSheet.Cells[RowNumber, 11] = product.FABSHORTAGE;
                    xlWorkSheet.Cells[RowNumber, 12] = product.FABCHKBY;

                    xlWorkSheet.Cells[RowNumber, 13] = product.PCACT;
                    xlWorkSheet.Cells[RowNumber, 14] = product.PCTIME;
                    xlWorkSheet.Cells[RowNumber, 15] = product.PCSHRT;
                    xlWorkSheet.Cells[RowNumber, 16] = product.PCCHKBY;

                    xlWorkSheet.Cells[RowNumber, 17] = product.HDACT;
                    xlWorkSheet.Cells[RowNumber, 18] = product.HDTIME;
                    xlWorkSheet.Cells[RowNumber, 19] = product.HDSHRT;
                    xlWorkSheet.Cells[RowNumber, 20] = product.HDCHKBY;

                    for(int i = 1; i <= 20; i++)
                    {
                        var oRange = ((Range)xlWorkSheet.Cells[RowNumber, i]);
                        oRange.Font.Size = 8;
                        oRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }
                    
                    decimal progress = ((count + 1) / (decimal)Data.Count) * 100;


                    SaveFileWorker.ReportProgress((int)progress);

                    RowNumber++;
                    count++;
                });

               

                xlWorkBook.SaveAs(FilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
                misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


                xlWorkBook.Close(false, misValue, misValue);
                

                
            }
            catch(Exception ex)
            {

            }

            xlApp.Quit();

            if(xlWorkSheet != null) Marshal.ReleaseComObject(xlWorkSheet);
            if (xlWorkBook != null)  Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return isSuccess;
        }
       

        public static List<EditableDataViewExcel> LoadDataFromExcelSheet(string FilePath,BackgroundWorker LoadingBackGroundWorker)
        {
            List<EditableDataViewExcel> editableDataViewExcels = new List<EditableDataViewExcel>();

            Microsoft.Office.Interop.Excel.Workbook xlWorkBookSource =  null;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetSource = null;
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                if (!string.IsNullOrEmpty(FilePath))
                {
                   


                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return editableDataViewExcels;
                    }


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
                            product.BRDNO = getCellsValue(xlWorkSheetSource, rowNumber, 1);
                            product.Raw = getCellsValue(xlWorkSheetSource, rowNumber, 2);
                            product.CATNO = getCellsValue(xlWorkSheetSource, rowNumber, 3);
                            product.NoOfBends = getCellsValue(xlWorkSheetSource, rowNumber, 4);
                            product.CatDesc = getCellsValue(xlWorkSheetSource, rowNumber, 5);
                            product.NetWtScrap = getCellsValue(xlWorkSheetSource, rowNumber, 6);
                            product.CompLoc = getCellsValue(xlWorkSheetSource, rowNumber, 7);
                            product.Qty = getCellsValue(xlWorkSheetSource, rowNumber, 8);

                            product.FABACT = getCellsValue(xlWorkSheetSource, rowNumber, 9); 
                            product.FABTIME = getCellsValue(xlWorkSheetSource, rowNumber, 10); 
                            product.FABSHORTAGE = getCellsValue(xlWorkSheetSource, rowNumber, 11);
                            product.FABCHKBY = getCellsValue(xlWorkSheetSource, rowNumber, 12);

                            product.PCACT = getCellsValue(xlWorkSheetSource, rowNumber, 13);
                            product.PCTIME = getCellsValue(xlWorkSheetSource, rowNumber, 14);
                            product.PCSHRT = getCellsValue(xlWorkSheetSource, rowNumber, 15);
                            product.PCCHKBY = getCellsValue(xlWorkSheetSource, rowNumber, 16);

                            product.HDACT = getCellsValue(xlWorkSheetSource, rowNumber, 17);
                            product.HDTIME = getCellsValue(xlWorkSheetSource, rowNumber, 18);
                            product.HDSHRT = getCellsValue(xlWorkSheetSource, rowNumber, 19);
                            product.HDCHKBY = getCellsValue(xlWorkSheetSource, rowNumber, 20);


                            if (
                                string.IsNullOrEmpty(product.BRDNO) ||
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

                }
            }
            catch (Exception e)
            {
                // Known Exception - Alternative Required

            }
            if(xlWorkSheetSource != null) Marshal.ReleaseComObject(xlWorkSheetSource);
            if (xlWorkBookSource != null) Marshal.ReleaseComObject(xlWorkBookSource);
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();

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
