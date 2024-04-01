using Microsoft.Maui.Controls.PlatformConfiguration;
using Syncfusion.Maui.DataGrid;
using Syncfusion.Maui.DataGrid.Exporting;
using System.Collections.ObjectModel;
namespace MauiApp1
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }

        private void ExportToExcel_Clicked(object sender, EventArgs e)
        {
            DataGridExcelExportingController excelExport = new DataGridExcelExportingController();
            DataGridExcelExportingOption options = new DataGridExcelExportingOption();

            // Export the selected rows
            // ObservableCollection<object> selectedItems = dataGrid.SelectedRows;
            // var excelEngine = excelExport.ExportToExcel(this.dataGrid, selectedItems);

            // Exclude the columns
            // var list = new List<string>();
            // list.Add("OrderID");
            // list.Add("CustomerID");
            // options.ExcludedColumns = list;

            // Exclude the header row
            //  options.CanExportHeader = false;

            // Start row and column index
            // options.StartRowIndex = 4;
            // options.StartColumnIndex = 2;

           // excelExport.RowExporting += ExcelExport_RowExporting; 
            var excelEngine = excelExport.ExportToExcel(this.dataGrid, options);
            var workbook = excelEngine.Excel.Workbooks[0];
            MemoryStream stream = new MemoryStream();
            workbook.SaveAs(stream);
            workbook.Close();
            excelEngine.Dispose();
            string OutputFilename = "DefaultDataGrid.xlsx";
            SaveService saveService = new();
            saveService.SaveAndView(OutputFilename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", stream);
        }

        private void ExcelExport_RowExporting(object? sender, DataGridRowExcelExportingEventArgs e)
        {
            if (!(e.Record.Data is OrderInfo))
                return;
            if (e.RowType == ExportRowType.RecordRow)
                e.Range.CellStyle.ColorIndex = Syncfusion.XlsIO.ExcelKnownColors.Aqua;
        }
    }
}
