﻿using A3DWinUtility;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A3DExcleUtility.Common.Classes
{
    public class ClsExcelWork
    {
        private static ClsExcelWork InstClsExcelWork = null;
        private ClsExcelWork()
        {

        }
        public static ClsExcelWork InstanceClsExcelWor
        {
            get
            {
                if (InstClsExcelWork == null)
                {

                    if (InstClsExcelWork == null)
                    {
                        InstClsExcelWork = new ClsExcelWork();
                    }

                }
                return InstClsExcelWork;
            }
        }
        public List<string> GetExcelSheet(string StrxmlFile)
        {
            try
            {
                List<string> LstExcelSheet = new List<string>();
                using (var workBook = new XLWorkbook(StrxmlFile))
                {
                    var results = workBook.Worksheets;
                    foreach (IXLWorksheet item in results)
                    {
                        LstExcelSheet.Add(item.Name);
                    }

                }
                return LstExcelSheet;
            }
            catch (Exception ex)
            {

                ClsMessage._IClsMessage.ProjectExceptionMessage(ex);
                return null;
            }

        }
        public void GetExcelColumns(string StrxmlFile, string StrExcelSheetName, DataTable DtReturnTable, string StrColumnName)
        {

            using (var workBook = new XLWorkbook(StrxmlFile))
            {
                var workSheet = workBook.Worksheet(StrExcelSheetName);
                var firstRowUsed = workSheet.FirstRowUsed();
                if (firstRowUsed == null) { return; }
                var firstPossibleAddress = workSheet.Row(firstRowUsed.RowNumber()).FirstCell().Address;
                var lastPossibleAddress = workSheet.LastCellUsed().Address;

                // Get a range with the remainder of the worksheet data (the range used)
                var range = workSheet.Range(firstPossibleAddress, lastPossibleAddress).AsRange(); //.RangeUsed();
                range.Clear(XLClearOptions.AllFormats);                                                                                // Treat the range as a table (to be able to use the column names)
                //var table = range.AsTable();
                IXLTable xLTable = null;
                if (range.Worksheet.Tables.Count() > 0)
                {
                    xLTable = range.Worksheet.Table(0);
                }
                else
                {
                    xLTable = range.CreateTable();
                }
                var table = xLTable;

                foreach (var item in table.Fields)
                {
                    DataRow dRow = DtReturnTable.NewRow();
                    dRow[StrColumnName] = item.Name;
                    DtReturnTable.Rows.Add(dRow);
                }

              
            }
        }

        public DataTable GetExcelData(string StrxmlFile, string StrExcelSheetName)
        {

            try
            {
                DataTable DtReturnTable = new DataTable();

                using (var workBook = new XLWorkbook(StrxmlFile))
                {
                    var workSheet = workBook.Worksheet(StrExcelSheetName);
                    var firstRowUsed = workSheet.FirstRowUsed();
                    if (firstRowUsed == null) { return null; }
                    var firstPossibleAddress = workSheet.Row(firstRowUsed.RowNumber()).FirstCell().Address;
                    var lastPossibleAddress = workSheet.LastCellUsed().Address;

                    // Get a range with the remainder of the worksheet data (the range used)
                    var range = workSheet.Range(firstPossibleAddress, lastPossibleAddress).AsRange(); //.RangeUsed();
                    range.Clear(XLClearOptions.AllFormats); // Treat the range as a table (to be able to use the column names)
                    //var table = range.AsTable();
                    IXLTable xLTable = null;
                    if (range.Worksheet.Tables.Count() > 0)
                    {
                        xLTable = range.Worksheet.Table(0);
                    }
                    else
                    {
                        xLTable = range.CreateTable();
                    }
                    var table = xLTable;
                    DtReturnTable = table.AsNativeDataTable();
                    return DtReturnTable;
                }
            }
            catch (Exception ex)
            {

                ClsMessage._IClsMessage.ProjectExceptionMessage(ex);
                return null;
            }
        }

        //private void GetDataFromExcel(string StrxmlFile)
        //{

        //    using (var workBook = new XLWorkbook(StrxmlFile))
        //    {
        //        var workSheet = workBook.Worksheet(1);
        //        var firstRowUsed = workSheet.FirstRowUsed();
        //        var firstPossibleAddress = workSheet.Row(firstRowUsed.RowNumber()).FirstCell().Address;
        //        var lastPossibleAddress = workSheet.LastCellUsed().Address;

        //        // Get a range with the remainder of the worksheet data (the range used)
        //        var range = workSheet.Range(firstPossibleAddress, lastPossibleAddress).AsRange(); //.RangeUsed();
        //                                                                                          // Treat the range as a table (to be able to use the column names)
        //        var table = range.AsTable();

        //        //Specify what are all the Columns you need to get from Excel
        //        var dataList = new List<string[]>
        //         {
        //             table.DataRange.Rows()
        //                 .Select(tableRow =>
        //                     tableRow.Field("Solution Number")
        //                         .GetString())
        //                 .ToArray(),
        //             table.DataRange.Rows()
        //                 .Select(tableRow => tableRow.Field("Name").GetString())
        //                 .ToArray(),
        //             table.DataRange.Rows()
        //             .Select(tableRow => tableRow.Field("Date").GetString())
        //             .ToArray()
        //         };
        //        //Convert List to DataTable
        //        var dataTable = ConvertListToDataTable(dataList);
        //        //To get unique column values, to avoid duplication
        //        var uniqueCols = dataTable.DefaultView.ToTable(true, "Solution Number");

        //        //Remove Empty Rows or any specify rows as per your requirement
        //        for (var i = uniqueCols.Rows.Count - 1; i >= 0; i--)
        //        {
        //            var dr = uniqueCols.Rows[i];
        //            if (dr != null && ((string)dr["Solution Number"] == "None" || (string)dr["Title"] == ""))
        //                dr.Delete();
        //        }
        //        Console.WriteLine("Total number of unique solution number in Excel : " + uniqueCols.Rows.Count);
        //    }
        //}
        private static DataTable ConvertListToDataTable(IReadOnlyList<string[]> list)
        {
            var table = new DataTable("CustomTable");
            var rows = list.Select(array => array.Length).Concat(new[] { 0 }).Max();

            table.Columns.Add("Solution Number");
            table.Columns.Add("Name");
            table.Columns.Add("Date");

            for (var j = 0; j < rows; j++)
            {
                var row = table.NewRow();
                row["Solution Number"] = list[0][j];
                row["Name"] = list[1][j];
                row["Date"] = list[2][j];
                table.Rows.Add(row);
            }
            return table;
        }



    }
}
