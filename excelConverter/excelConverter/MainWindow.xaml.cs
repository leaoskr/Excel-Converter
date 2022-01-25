using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace excelConverter
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
            //单个文件
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel File (*.xlsx)|*.xlsx;*.xlsm|All Files (*.*)|*.*";
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            
            if (dialog.ShowDialog() == true)
            {

                string filePath = dialog.FileName;

                FileInfo fileInfo = new FileInfo(filePath);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(fileInfo))
                {
                    //ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        worksheet.DeleteRow(1, 8);
                        worksheet.DeleteColumn(1);
                        //worksheet.DeleteRow(1, 1);//测试
                        //worksheet.DeleteColumn(1);//测试

                        // get number of rows in the sheet
                        int rows = worksheet.Dimension.Rows;
                        worksheet.DeleteRow(rows - 4, rows);


                        // loop through the worksheet rows
                        for (int i = 1; i <= rows; i++)
                        {
                            for (int j = 1; j <= rows; j++)
                            {
                                if (worksheet.Cells[i, j].Value != null)
                                {
                                    if (worksheet.Cells[i, j].Merge) //if the cell is merged
                                    {
                                        var mergeId = worksheet.MergedCells[i, j];//merge cell range
                                        var temp = worksheet.Cells[mergeId].First().Value.ToString(); // merge cell value

                                        //clear the merge cell range
                                        worksheet.Select(mergeId);
                                        worksheet.SelectedRange.Clear();

                                        //give value to each cell
                                        worksheet.Cells[mergeId].Value = temp;

                                    }

                                }

                            }

                        }
                    }
                    
                    

                    // save changes
                    package.Save();//save to a new location: pack.SaveAs(new FileInfo(outputFilePath));

                    executeNotify.Text = "finished!!";
                }

            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            
            //多个文件
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel File (*.xlsx)|*.xlsx;*.xlsm|All Files (*.*)|*.*";
            dialog.Multiselect = true;
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (dialog.ShowDialog() == true)
            {
                var files = dialog.FileNames;

                var resultFile = @"C:\Users\fston\Desktop\result.xlsx";
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var masterPackage = new ExcelPackage(new FileInfo(@"C:\Users\fston\Desktop\first.xlsx")))
                {
                    int count = 0;

                    foreach (var file in files)
                    {
                        using (var pckg = new ExcelPackage(new FileInfo(file)))
                        {
                            
                            foreach (var sheet in pckg.Workbook.Worksheets)
                            {
                                //check name of worksheet, in case that worksheet with same name already exist exception will be thrown by EPPlus

                                string workSheetName = sheet.Name;

                                foreach (var masterSheet in masterPackage.Workbook.Worksheets)
                                {
                                    if (sheet.Name == masterSheet.Name)
                                    {
                                        workSheetName = string.Format("{0}_{1}_{2}", count++, workSheetName, DateTime.Now.ToString("yyyyMMddhhssmmm"));
                                    }
                                }

                                //add new sheet
                                masterPackage.Workbook.Worksheets.Add(workSheetName, sheet);
                            }
                        }

                    }
                    masterPackage.SaveAs(new FileInfo(resultFile));

                    combineNotify.Text = "finished!!";
                }

            }


            
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel File (*.xlsx)|*.xlsx;*.xlsm|All Files (*.*)|*.*";
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            if (dialog.ShowDialog() == true)
            {

                string filePath = dialog.FileName;

                FileInfo fileInfo = new FileInfo(filePath);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(fileInfo))
                {
                    var sheet = package.Workbook.Worksheets;
                    int count = package.Workbook.Worksheets.Count;


                    for (int i = 1; i < count; i++)
                    {

                        int src_col = sheet[i].Dimension.Columns;
                        int src_row = sheet[i].Dimension.Rows;

                        int des_row = sheet[0].Dimension.Rows;

                        sheet[i].Cells[1, 1, src_row, src_col].Copy(sheet[0].Cells[des_row + 1, 1]);
                        
                    }

                    for (int j = count-1; j > 0; j--)
                    {
                        package.Workbook.Worksheets.Delete(sheet[j]);
                    }

                    package.Save();
                    mergeNotify.Text = "finished!!";
                }

                
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            combineNotify.Text = "";
            executeNotify.Text = "";
            mergeNotify.Text = "";
        }
    }
}
