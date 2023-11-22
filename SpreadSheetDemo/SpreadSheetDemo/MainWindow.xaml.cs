using Syncfusion.UI.Xaml.Spreadsheet;
using Syncfusion.UI.Xaml.Spreadsheet.GraphicCells;
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
using Syncfusion.UI.Xaml.SpreadsheetHelper;

namespace SpreadSheetDemo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            //creating 6 sheets to the spreadsheet
            //spreadsheet.Create(6);

            //opening an existing excel file using SfSpreadSheet control
            //spreadsheet.Open(@"D:\book1.xlsx");

            ////Saving the file after editing
            //spreadsheet.Save();

            //For importing charts,
            //spreadsheet.AddGraphicChartCellRenderer(new GraphicChartCellRenderer());

        }
    }
}