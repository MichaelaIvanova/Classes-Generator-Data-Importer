using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

public static class DataTableHelper
{
    //Convert Collection to DataTable (Not including List Properties)
    public static DataTable ToDataTable<T>(IEnumerable<T> collection)
    {
        Type t = typeof(T);

        string tableName = GetTableName(t);

        DataTable dt = new DataTable(tableName);
        List<PropertyInfo> properties = t.GetProperties()
                                        .Where(p => !p.PropertyType.FullName.Contains("Generic.List"))
                                        .OrderBy(o => o.GetCustomAttributes(false).OfType<DisplayColumnAttribute>().First().DisplayOrder)
                                        .ToList();

        //Inspect the properties and create the columns in the DataTable
        foreach (PropertyInfo pi in properties)
        {
            string columnName = GetColumnName(pi);

            Type columnType = pi.PropertyType;
            if ((columnType.IsGenericType))
            {
                columnType = columnType.GetGenericArguments()[0];
            }
            dt.Columns.Add(columnName, columnType);
        }

        //Populate the data table
        foreach (T item in collection)
        {
            DataRow dr = dt.NewRow();
            dr.BeginEdit();
            foreach (PropertyInfo pi in properties)
            {
                string columnName = GetColumnName(pi);

                if (pi.GetValue(item, null) != null)
                {
                    dr[columnName] = pi.GetValue(item, null);
                }
            }
            dr.EndEdit();
            dt.Rows.Add(dr);
        }
        return dt;
    }

    private static string GetColumnName(PropertyInfo pi)
    {
        string columnName = pi.Name.ToUpper();

        var columnAttr = (DisplayColumnAttribute)pi.GetCustomAttributes(false)
                            .FirstOrDefault(x => x.GetType() == typeof(DisplayColumnAttribute));

        if (columnAttr != null)
        {
            columnName = columnAttr.DisplayName;
        }

        return columnName;
    }

    private static string GetTableName(Type t)
    {
        var tableAttr = (DisplayTableAttribute)t.GetCustomAttributes(false)
                        .FirstOrDefault(x => x.GetType() == typeof(DisplayTableAttribute));

        string tableName = "Table";
        if (tableAttr != null)
        {
            tableName = tableAttr.Display;
        }

        return tableName;
    }
}
