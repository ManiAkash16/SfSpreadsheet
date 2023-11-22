using Syncfusion.UI.Xaml.CellGrid;
using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Syncfusion.UI.Xaml.Grid.ScrollAxis;
using Syncfusion.UI.Xaml.Spreadsheet;
using Syncfusion.UI.Xaml.Spreadsheet.Helpers;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
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

namespace SpreadSheetWorking
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            spreadsheet.WorksheetAdded += spreadsheet_WorksheetAdded;
            spreadsheet.WorksheetRemoved += spreadsheet_WorksheetRemoved;
            spreadsheet.WorkbookLoaded += spreadsheet_WorkbookLoaded;
            spreadsheet.WorkbookUnloaded += spreadsheet_WorkbookUnloaded;
            spreadsheet.PropertyChanged += spreadsheet_PropertyChanged;
            spreadsheet.Open(@"D:\Attendence\AttendanceReport_October2023.xlsx");

            //Accessing a worksheet

            //By Specifying the sheet name as,
            //var sheet2 = spreadsheet.Workbook.Worksheets["Sheet3"];
            //sheet2.Activate();
            //Access the Active worksheet as,
            //var sheet3 = spreadsheet.ActiveSheet;


        }

        //Accessing the grid

        void spreadsheet_WorksheetAdded(object sender, WorksheetAddedEventArgs args)
        {

            //Access the Active SpreadsheetGrid and hook the events associated with it.
            var grid = spreadsheet.ActiveGrid;
            grid.CurrentCellActivated += grid_CurrentCellActivated;
            //ClearCell();
        }

        void spreadsheet_WorksheetRemoved(object sender, WorksheetRemovedEventArgs args)
        {

            //Access the Active SpreadsheetGrid and unhook the events associated with it
            var grid = spreadsheet.ActiveGrid;
            grid.CurrentCellActivated -= grid_CurrentCellActivated;
        }

        void grid_CurrentCellActivated(object sender, CurrentCellActivatedEventArgs args)
        {
            var cell = spreadsheet.CurrentCellValue;
            MessageBox.Show(cell);
            //RefreshCell();
            //ModifyCheck();
            spreadsheet.ActiveGrid.ShowHidePopup(true);
        }


        //by event

        void spreadsheet_WorkbookLoaded(object sender, WorkbookLoadedEventArgs args)
        {
            //var sheet = spreadsheet.ActiveSheet;
            //spreadsheet.GridCollection[sheet.Name].RowCount=5;
            //spreadsheet.GridCollection[sheet.Name].ColumnCount = 12;

            //setting active sheet

            //spreadsheet.SetActiveSheet("Sheet3");
            //ValueChanged();
            
            var grid = spreadsheet.ActiveGrid;
            grid.CurrentCellActivated += grid_CurrentCellActivated;
            AccessingCell();
        }

        void spreadsheet_WorkbookUnloaded(object sender, WorkbookUnloadedEventArgs args)
        {

            // Access a cell value by using "Value" Property
            var cellValue = spreadsheet.Workbook.Worksheets[0].Range["B3"].Value;

            // Access a cell value by using "DisplayText" Property. 
            var displayValue = spreadsheet.Workbook.Worksheets[1].Range[4, 1].DisplayText;
            var displayValu3e = spreadsheet.Workbook.Worksheets[1].Range[4, 1].Value;
            var grid = spreadsheet.ActiveGrid;

        }

        void spreadsheet_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName =="ActiveSheet")
            {
                var name = e.PropertyName;
            }
        }

        void AccessingCell()
        {
            //Accessing a cell

            //// Access a cell by specifying cell address.
            //var cell = spreadsheet.Workbook.Worksheets[0].Range["A3"];

            //// Access a cell by specifying cell row and column index. 
            //var cell1 = spreadsheet.Workbook.Worksheets[0].Range[3, 1];

            //// Access a cells by specifying user defined name.
            //var cell2 = spreadsheet.Workbook.Worksheets[0].Range["Namerange"];

            // Accessing a range of cells by specifying cell's address.
            var cell3 = spreadsheet.Workbook.Worksheets[0].Range["A5:C8"];
            cell3.Value = "HI";

            //// Accessing a range of cells specifying cell row and column index.
            //var cell4 = spreadsheet.Workbook.Worksheets[0].Range[15, 1, 15, 3];
        }

        void ValueChanged()
        {
            //Setting the value or formula to a cell
            var range = spreadsheet.ActiveSheet.Range[2, 2];
            var range1 = spreadsheet.ActiveSheet.Range[2, 4];
            var range3 = spreadsheet.ActiveSheet.Range[2, 3];
            spreadsheet.ActiveGrid.InvalidateCell(2, 2);
            spreadsheet.ActiveGrid.SetCellValue(range, "cellValue");
            spreadsheet.ActiveGrid.SetCellValue(range3, "hello");
            ExcelFormatType x = range1.FormatType;

            range3.AddComment();
        }
        void ClearCell()
        {
            //To clear the contents in the range alone,
            spreadsheet.Workbook.Worksheets[0].Range[2, 4].Clear(true);

            //To clear the contents along with its formatting in the range,   
            spreadsheet.Workbook.Worksheets[0].Range[2, 2].Clear();

            //To clear the range with specified ExcelClearOptions,
            spreadsheet.Workbook.Worksheets[0].Range[2, 3].Clear(ExcelClearOptions.ClearContent);
        }

        void RefreshCell()
        {
            ////Invalidates the mentioned cell in the grid,
            //spreadsheet.ActiveGrid.InvalidateCell(3, 3);

            //Invalidates the range ,
            //var range = GridRangeInfo.Cells(5, 4, 6, 7);
            //spreadsheet.ActiveGrid.InvalidateCell(range);

            //Invalidates all the cells in the grid,
            spreadsheet.ActiveGrid.InvalidateCells();

            ////Invalidates the measurement state(layout) of grid,
            //spreadsheet.ActiveGrid.InvalidateVisual();

            ////Invalidates the cell borders in the range,
            //var range1 = GridRangeInfo.Cells(2, 4, 6, 4);
            //spreadsheet.ActiveGrid.InvalidateCellBorders(range1);
        }

        void ModifyCheck()
        {
            var workbook = spreadsheet.Workbook as WorkbookImpl;
            BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static;
            var value = typeof(WorkbookImpl).GetProperty("IsCellModified", bindFlags).GetValue(workbook);
        }
    }
}