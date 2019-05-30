using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using OfficeOpenXml;

public class XlsxReader
{
    /// <summary>
    /// 将指定Excel文件的内容读取到DataSet中
    /// </summary>
    public static DataSet ReadXlsxFile(string filePath, out string errorString)
    {
        errorString = "";
        FileInfo existingFile = new FileInfo(filePath);
        try
        {
            ExcelPackage package = new ExcelPackage(existingFile);
            int vSheetCount = package.Workbook.Worksheets.Count;
            //Console.WriteLine("sheet count: {0}", vSheetCount);

            // 必须存在数据表
            bool isFoundDateSheet = false;
            var sheetNameData = AppValues.EXCEL_DATA_SHEET_NAME.Replace("$", "");
            // 可选配置表
            bool isFoundConfigSheet = false;
            var sheetNameConfig = AppValues.EXCEL_CONFIG_SHEET_NAME.Replace("$", "");

            foreach (var s in package.Workbook.Worksheets)
            {
                //Console.WriteLine(s.Name);
                var sheetName = s.Name;
                if (sheetName == sheetNameData)
                    isFoundDateSheet = true;
                else if (sheetName == sheetNameConfig)
                    isFoundConfigSheet = true;
            }
            if (!isFoundDateSheet)
            {
                errorString = string.Format("错误：{0}中不含有Sheet名为{1}的数据表", filePath, sheetNameData);
                return null;
            }

            var ds = new DataSet();
            foreach (var s in package.Workbook.Worksheets)
            {
                DataTable dt = ExcelWorksheetToDataTable(s);
                dt.TableName = s.Name + "$";
                ds.Tables.Add(dt);
            }
            return ds;
        }
        catch (Exception err)
        {
            errorString = err.Message + "\n" + err.StackTrace;
        }
        return null;
    }

    private static DataTable ExcelWorksheetToDataTable(ExcelWorksheet worksheet)
    {
        DataTable dt = new DataTable();

        //check if the worksheet is completely empty
        if (worksheet.Dimension == null)
        {
            return dt;
        }

        //create a list to hold the column names
        List<string> columnNames = new List<string>();

        //needed to keep track of empty column headers
        int currentColumn = 1;

        for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
        {
            var cell = worksheet.Cells[1, i];

            string columnName = cell.Text.Trim();
            //check if the previous header was empty and add it if it was
            if (cell.Start.Column != currentColumn)
            {
                for (; currentColumn < cell.Start.Column; currentColumn++)
                {
                    columnNames.Add("Header_" + currentColumn);
                    dt.Columns.Add("Header_" + currentColumn);
                }
                currentColumn = cell.Start.Column;
            }

            //add the column name to the list to count the duplicates
            columnNames.Add(columnName);

            //count the duplicate column names and make them unique to avoid the exception
            //A column named 'Name' already belongs to this DataTable
            int occurrences = columnNames.FindAll(x => x.Equals(columnName)).Count;
            if (occurrences > 1)
            {
                columnName = columnName + "_" + occurrences;
            }

            //add the column to the datatable
            dt.Columns.Add(columnName);

            currentColumn++;
        }

        //start adding the contents of the excel file to the datatable
        for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
        {
            var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
            var cols = row.Columns;
            DataRow newRow = dt.NewRow();

            //loop all cells in the row
            foreach (var cell in row)
            {
                var col = cell.Start.Column - 1;
                if (newRow.ItemArray.Length > col)
                    newRow[col] = cell.Text;
            }

            dt.Rows.Add(newRow);
        }

        if (dt.Rows.Count > 1)
        {
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                var row = dt.Rows[i];
                bool IsNull = true;
                foreach (var item in row.ItemArray)
                {
                    IsNull &= string.IsNullOrEmpty(item.ToString().Trim());
                }
                if (IsNull)
                {
                    dt.Rows.Remove(row);
                }
                else
                {
                    break;
                }
            }
        }

        return dt;
    }
}
