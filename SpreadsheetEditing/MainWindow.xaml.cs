using System;
using System.Collections.Generic;
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
using Syncfusion.UI.Xaml.Spreadsheet;
using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Syncfusion.UI.Xaml.Spreadsheet.Helpers;
using Syncfusion.XlsIO;

namespace SpreadsheetEditing
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public int count=0, k=1;
        public MainWindow()
        {
            InitializeComponent();
            spreadsheet.Open(@"D:\Attendence\AttendanceReport_October2023.xlsx");
            spreadsheet.WorkbookLoaded += spreadsheet_WorkbookLoaded;
        }

        private void ActiveGrid_CurrentCellBeginEdit(object sender, CurrentCellBeginEditEventArgs e)
        {

            spreadsheet.ActiveSheet.Range["F5"].Value = "Hello";
        }

        private void grid_CurrentCellActivated(object sender, CurrentCellActivatedEventArgs e)
        {

            //spreadsheet.ActiveGrid.CurrentCell.EndEdit();

        }

        private void spreadsheet_WorkbookLoaded(object sender, WorkbookLoadedEventArgs args)
        {
            if(count != 0)
            {
                spreadsheet.ActiveGrid.AllowEditing = true;
                var grid = spreadsheet.ActiveGrid;
                var sheet = spreadsheet.ActiveSheet;
                
                var excelStyle = sheet.Range["A2"].CellStyle;

                //To unlock a cell,           
                //excelStyle.Locked = false;

                //To lock a cell, 
                excelStyle.Locked = true;
                grid.CurrentCellActivated += grid_CurrentCellActivated;
                grid.CurrentCellBeginEdit += ActiveGrid_CurrentCellBeginEdit;
                grid.CurrentCellEndEdit += Grid_CurrentCellEndEdit;
                DataValidation();
            }
            count++;
        }

        private void Grid_CurrentCellEndEdit(object sender, CurrentCellEndEditEventArgs e)
        {
            spreadsheet.ActiveGrid.CurrentCellActivated -= grid_CurrentCellActivated;
        }

        private void DataValidation()
        {
            //Number Validation
            IDataValidation validation = spreadsheet.ActiveSheet.Range["A5"].DataValidation;
            validation.AllowType = ExcelDataType.Integer;
            validation.CompareOperator = ExcelDataValidationComparisonOperator.Between;
            validation.FirstFormula = "4";
            validation.SecondFormula = "15";
            validation.ShowErrorBox = true;
            validation.ErrorBoxText = "Accepts values only between 4 to 15";

            //Date Validation
            IDataValidation validation1 = spreadsheet.ActiveSheet.Range["B4"].DataValidation;
            validation1.AllowType = ExcelDataType.Date;
            validation1.CompareOperator = ExcelDataValidationComparisonOperator.Greater;
            validation1.FirstDateTime = new DateTime(2016, 5, 5);
            validation1.ShowErrorBox = true;
            validation1.ErrorBoxText = "Enter the date value which is greater than 05/05/2016";

            //TextLength Validation
            IDataValidation validation2 = spreadsheet.ActiveSheet.Range["A3:B3"].DataValidation;
            validation2.AllowType = ExcelDataType.TextLength;
            validation2.CompareOperator = ExcelDataValidationComparisonOperator.LessOrEqual;
            validation2.FirstFormula = "4";
            validation2.ShowErrorBox = true;
            validation2.ErrorBoxText = "Text length should be lesser than or equal 4 characters";

            //List Validation
            IDataValidation validation3 = spreadsheet.ActiveSheet.Range["D4"].DataValidation;
            validation3.ListOfValues = new string[] { "10", "20", "30" };

            //Custom Validation
            IDataValidation validation4 = spreadsheet.ActiveSheet.Range["E4"].DataValidation;
            validation4.AllowType = ExcelDataType.Formula;
            validation4.FirstFormula = "=A1+A2>0";
            validation4.ErrorBoxText = "Sum of the values in A1 and A2 should be greater than zero";
        }
    }
}
