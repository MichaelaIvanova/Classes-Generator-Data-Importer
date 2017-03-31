using System.Data;
using System.IO;
using ExcelLibrary;
using OfficeOpenXml;

public static class ExcelHelper
{
    public static DataTable ExcelToDataTable(string excelPath)
    {
        DataSet ds = DataSetHelper.CreateDataSet(excelPath);
        return ds.Tables[0];
    }

    public static DataSet ExcelToDataSet(string excelPath)
    {
        return DataSetHelper.CreateDataSet(excelPath);
    }

    public static void DataTableToExcel(DataTable dt, string path)
    {
        DataSet ds = new DataSet();
        ds.Tables.Add(dt);
        DataSetHelper.CreateWorkbook(path, ds);
    }

    public static void DataSetToExcel(DataSet ds, string path)
    {
        //DataSetHelper.CreateWorkbook(path, ds);
        FileInfo file = new FileInfo(path);

        using (ExcelPackage pck = new ExcelPackage(file))
        {
            foreach (DataTable dataTable in ds.Tables)
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add(dataTable.TableName);
                ws.Cells["A1"].LoadFromDataTable(dataTable, true);
                var range = ExcelRange.GetAddress(1, 1, dataTable.Rows.Count+1, dataTable.Columns.Count);
                pck.Workbook.Names.Add(dataTable.TableName, ws.Cells[range]);
            }
            pck.Save();    
        }
    }
}
