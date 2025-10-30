# Exporting-DataGrid-to-Excel-Made-Easy-in-.NET-MAUI
Exporting a Syncfusion .NET MAUI DataGrid (SfDataGrid) to Excel is simple and flexible. This guide explains how the included sample exports grid data to an .xlsx file, how to configure export options, and how to extend the behavior with styling and selection-based export.

For more details, please refer the official documentation: [Export To Excel in MAUI DataGrid (SfDataGrid)](https://help.syncfusion.com/maui/datagrid/export-to-excel)

## xaml
```
<ContentPage.BindingContext>
    <local:OrderInfoRepository x:Name="viewModel" />
</ContentPage.BindingContext>
<ContentPage.Content>
    <StackLayout>
            <Button Text="Export To Excel" WidthRequest="200" HeightRequest="50"  
                Clicked="ExportToExcel_Clicked" />
        <syncfusion:SfDataGrid x:Name="dataGrid"
                            Margin="20"
                            VerticalOptions="FillAndExpand"
                            ItemsSource="{Binding OrderInfoCollection}"
                            GridLinesVisibility="Both"
                            HeaderGridLinesVisibility="Both"
                            AutoGenerateColumnsMode="None"
                            SelectionMode="Multiple"
                            ColumnWidthMode="Auto">
            <syncfusion:SfDataGrid.Columns>
                <syncfusion:DataGridNumericColumn Format="D"
                                                HeaderText="Order ID"
                                                MappingName="OrderID">
                </syncfusion:DataGridNumericColumn>
                <syncfusion:DataGridTextColumn HeaderText="Customer ID"
                                            MappingName="CustomerID">
                </syncfusion:DataGridTextColumn>
                <syncfusion:DataGridTextColumn MappingName="Customer"
                                            HeaderText="Customer">
                </syncfusion:DataGridTextColumn>
                <syncfusion:DataGridTextColumn HeaderText="Ship City"
                                            MappingName="ShipCity">
                </syncfusion:DataGridTextColumn>
                <syncfusion:DataGridTextColumn HeaderText="Ship Country"
                                            MappingName="ShipCountry">
                </syncfusion:DataGridTextColumn>
            </syncfusion:SfDataGrid.Columns>
        </syncfusion:SfDataGrid>
    </StackLayout>
</ContentPage.Content>
```

## C#
The code-behind wires the button click to export the grid via DataGridExcelExportingController. You can export all rows, only selected rows, exclude certain columns, and even style rows during export using the RowExporting event.
```
private void ExportToExcel_Clicked(object sender, EventArgs e)
{
    DataGridExcelExportingController excelExport = new DataGridExcelExportingController();
    DataGridExcelExportingOption options = new DataGridExcelExportingOption();

    // Export the selected rows (optional)
    // ObservableCollection<object> selectedItems = dataGrid.SelectedRows;
    // var excelEngine = excelExport.ExportToExcel(this.dataGrid, selectedItems);

    // Exclude columns from export (optional)
    // options.ExcludedColumns = new List<string> { "OrderID", "CustomerID" };

    // Exclude header row (optional)
    // options.CanExportHeader = false;

    // Start row/column index in the worksheet (optional)
    // options.StartRowIndex = 4;
    // options.StartColumnIndex = 2;

    // Customize rows while exporting (optional)
    // excelExport.RowExporting += ExcelExport_RowExporting;

    var excelEngine = excelExport.ExportToExcel(this.dataGrid, options);
    var workbook = excelEngine.Excel.Workbooks[0];
    using MemoryStream stream = new MemoryStream();
    workbook.SaveAs(stream);
    workbook.Close();
    excelEngine.Dispose();

    string outputFilename = "DefaultDataGrid.xlsx";
    SaveService saveService = new();
    saveService.SaveAndView(outputFilename,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        stream);
}

private void ExcelExport_RowExporting(object? sender, DataGridRowExcelExportingEventArgs e)
{
    if (e.RowType == ExportRowType.RecordRow)
    {
        // Example: apply a background color to record rows in Excel
        e.Range.CellStyle.ColorIndex = Syncfusion.XlsIO.ExcelKnownColors.Aqua;
    }
}
```

##### Conclusion
 
I hope you enjoyed learning about how to export DataGrid to Excel in .NET MAUI DataGrid (SfDataGrid).
 
You can refer to our [.NET MAUI DataGrid’s feature tour](https://www.syncfusion.com/maui-controls/maui-datagrid) page to learn about its other groundbreaking feature representations. You can also explore our [.NET MAUI DataGrid Documentation](https://help.syncfusion.com/maui/datagrid/getting-started) to understand how to present and manipulate data. 
For current customers, you can check out our .NET MAUI components on the [License and Downloads](https://www.syncfusion.com/sales/teamlicense) page. If you are new to Syncfusion, you can try our 30-day [free trial](https://www.syncfusion.com/downloads/maui) to explore our .NET MAUI DataGrid and other .NET MAUI components.
 
If you have any queries or require clarifications, please let us know in the comments below. You can also contact us through our [support forums](https://www.syncfusion.com/forums), [Direct-Trac](https://support.syncfusion.com/create) or [feedback portal](https://www.syncfusion.com/feedback/maui?control=sfdatagrid), or the feedback portal. We are always happy to assist you!
