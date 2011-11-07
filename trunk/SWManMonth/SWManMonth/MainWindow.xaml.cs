using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace SWManMonth
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void pbBrowseFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory = @"C:\";
            openFile.Filter = "Excel 97-2003 (.xls)|*.xls|Excel 2007 (.xlsx)|*.xlsx";
            openFile.FilterIndex = 2;

            if (openFile.ShowDialog() == true)
            {
                tbBrowseFile.Text = openFile.FileName;
            }
        }

        private void pbProcessFile_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = new Excel.Application(); ;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;
            int countWorkSheets = 0;
            int rCount = 0;
            int cCount = 0;
            String text = null;

            // for debug

            String wsName = null;
            int linha = 0;
            int coluna = 0;

            try
            {
                xlWorkBook = xlApp.Workbooks.Open(tbBrowseFile.Text, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                List<String> projectList = new List<String>();
                List<Model> modelList = new List<Model>();

                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    xlRange = ws.UsedRange;
                    wsName = ws.Name;
                    ExcelColumns ec = new ExcelColumns(xlRange.Columns.Count);
                    String[] head = ec.Columns;
                    String cell = null;

                    linha = coluna = 0;

                    for (cCount = 1; cCount < xlRange.Columns.Count; cCount++)
                    {
                        cell = head[cCount].ToString() + "1";
                        text = (string)ws.get_Range(cell, cell).Cells.Value2;
                        if (text == null)
                        {
                            text = "";
                            continue;
                        }
                        coluna++;
                        if (projectList!=null)
                        {
                            Model m = new Model();
                            m.ModelCode = text;
                            CA ca = new CA();
                            cell = head[cCount].ToString() + "2";
                            text = (string)ws.get_Range(cell, cell).Cells.Value2;
                            ca.Country = text;
                            cell = head[cCount].ToString() + "3";
                            text = (string)ws.get_Range(cell, cell).Cells.Value2;
                            ca.CarrierName = text;

                            linha = xlRange.Rows.Count;

                            cell = head[cCount].ToString() + xlRange.Rows.Count.ToString();
                            Object obj = ws.get_Range(cell, cell).Cells.Value2;
                            double tmp = 0.0;
                            if (obj.GetType() == typeof(string))
                                tmp = Double.Parse((string)obj);
                            else if (obj.GetType() == typeof(double))
                                tmp = (double)obj;

//                            ca.MediumManMonth = tmp / 1000;
                            ca.MediumManMonth = tmp;

                            if (modelList.Count > 0)
                            {
                                Model tempModel = modelList.Find(delegate(Model mm) { return (mm.ModelCode == m.ModelCode && mm.ModelCAs.CarrierName==ca.CarrierName && mm.ModelCAs.Country==ca.Country && mm.ModelCAs.MediumManMonth>0);});
                                if (tempModel == null)
                                {
                                    m.ModelCAs = ca;
                                    modelList.Add(m);
                                }
                                else
                                {
                                    ca.MediumManMonth += tempModel.ModelCAs.MediumManMonth;
                                    modelList.Remove(tempModel);
                                    tempModel.ModelCAs = ca;
                                    modelList.Add(tempModel);
                                }
                            }
                            else
                            {
                                m.ModelCAs = ca;
                                modelList.Add(m);
                                int pos = 0;
                                pos = modelList.IndexOf(m);
                                pos = modelList.LastIndexOf(m);
                            }
                        }
                    }
                }

                xlWorkBook.Close();
                xlWorkBook = null;
                xlWorkBook = xlApp.Workbooks.Open(tbDestFile.Text, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Sheets xlWorkSheets = (Excel.Sheets)xlWorkBook.Worksheets;
                xlWorkSheet = xlWorkSheets.get_Item(1);

                xlRange = xlWorkSheet.get_Range("A1", "C" + modelList.Count.ToString());
                rCount = 1;
                foreach(Model m in modelList)
                {
                    xlRange = xlWorkSheet.get_Range("A" + rCount.ToString(), "A" + rCount.ToString());
                    xlRange.Value = m.ModelCode.ToString();
                    xlRange = xlWorkSheet.get_Range("B" + rCount.ToString(), "B" + rCount.ToString());
                    xlRange.Value = m.ModelCAs.Country.ToString();
                    xlRange = xlWorkSheet.get_Range("C" + rCount.ToString(), "C" + rCount.ToString());
                    xlRange.Value = m.ModelCAs.CarrierName.ToString();
                    xlRange = xlWorkSheet.get_Range("D" + rCount.ToString(), "D" + rCount.ToString());
                    xlRange.NumberFormat = "0.00000";
                    xlRange.Value = m.ModelCAs.MediumManMonth;

                    rCount++;

                }

                xlWorkBook.Close();
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                String msg = "Planihla: " + wsName + "\nlinha: " + linha + "\ncoluna: " + coluna + "\n\n";
                MessageBox.Show(msg + ex.Message);
            }
            finally
            {
                xlApp.Quit();
            }
        }

        private void pbBrowseDestFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.InitialDirectory = @"C:\";
            saveFile.Filter = "Excel 97-2003 (.xls)|*.xls|Excel 2007 (.xlsx)|*.xlsx";
            saveFile.FilterIndex = 2;

            if (saveFile.ShowDialog() == true)
                tbDestFile.Text = saveFile.FileName;

        }
    }
}
