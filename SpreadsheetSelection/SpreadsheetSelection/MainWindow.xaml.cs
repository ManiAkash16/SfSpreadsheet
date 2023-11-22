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
using Syncfusion.UI.Xaml.CellGrid;
using Syncfusion.UI.Xaml.Spreadsheet;
using Syncfusion.UI.Xaml.Spreadsheet.Commands;
using Syncfusion.UI.Xaml.Spreadsheet.Helpers;
using Syncfusion.XlsIO;

namespace SpreadsheetSelection
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            spreadsheet.WorkbookLoaded += spreadsheet_WorkbookLoaded;
            spreadsheet.Create(4);
        }

        private void spreadsheet_WorkbookLoaded(object sender,WorkbookLoadedEventArgs args)
        {
            spreadsheet.ActiveGrid.AllowSelection = true;

            //accessing the selected range

            //var rangeList = spreadsheet.ActiveGrid.SelectedRanges;

            //accessing the current cell

            //var cell = spreadsheet.ActiveGrid.SelectionController.CurrentCell;
            //AddOrClear();
            //MoveCurrentCell();
            GridToExcelRange();
        }

        void AddOrClear()
        {

            //adding or clearing the selection

            //To Add the Selection for range,
            //spreadsheet.ActiveGrid.SelectionController.AddSelection(GridRangeInfo.Cells(4, 6, 5, 8));

            //To Add the Selection for particular row,
            //spreadsheet.ActiveGrid.SelectionController.AddSelection(GridRangeInfo.Row(4));

            //To Add the Selection for multiple rows,
            //spreadsheet.ActiveGrid.SelectionController.AddSelection(GridRangeInfo.Rows(4, 9));

            //To Add the Selection for particular column,
            //spreadsheet.ActiveGrid.SelectionController.AddSelection(GridRangeInfo.Col(5));

            //To Add the Selection for multiple columns,
            spreadsheet.ActiveGrid.SelectionController.AddSelection(GridRangeInfo.Cols(5, 10));

            //To Clear the Selection,
            spreadsheet.ActiveGrid.SelectionController.ClearSelection();
        }
        
        void MoveCurrentCell()
        {
            //Moves current cell to the mentioned row and column index of cell,
            //spreadsheet.ActiveGrid.CurrentCell.MoveCurrentCell(5, 5);

            //For moving the current cell to a different sheet,
            spreadsheet.SetActiveSheet("Sheet2");
            spreadsheet.ActiveGrid.CurrentCell.MoveCurrentCell(6, 5);
        }

        void GridToExcelRange()
        {
            //converting GridRangeInfo into IRange
            spreadsheet.Workbook.Worksheets[0].Range[4, 5].Value = "hello";
            var excelRange = GridExcelHelper.ConvertGridRangeToExcelRange(GridRangeInfo.Cell(4, 5), spreadsheet.ActiveGrid);
            spreadsheet.Workbook.Worksheets[0].Range[4,5].Value = excelRange;
        }
}
}
