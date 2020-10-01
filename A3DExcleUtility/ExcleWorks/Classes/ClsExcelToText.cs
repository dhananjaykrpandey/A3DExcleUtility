using A3DExcleUtility.Common.Classes;
using A3DWinUtility;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A3DExcleUtility.ExcleWorks.Classes
{
    class ClsExcelToText
    {
        private static ClsExcelToText _iClsExcelToText = null;
        public ClsExcelToText()

        {

        }
        public static ClsExcelToText _IClsExcelToText
        {
            get
            {
                if (_iClsExcelToText == null)
                {
                    _iClsExcelToText = new ClsExcelToText();
                }
                return _iClsExcelToText;
            }

        }

        public List<string> GetExcelSheet(string StrExcelFileName)
        {
            try
            {
                return ClsExcelWork.InstanceClsExcelWor.GetExcelSheet(StrExcelFileName);
            }
            catch (Exception)
            {

                throw;
            }
        }
        public DataTable GetExcelColumns(string StrExcelFile, string StrExcelSheetName)
        {
            try
            {
                DataTable DtExcelCol = new DataTable();
                DtExcelCol.Columns.AddRange(new DataColumn[] { new DataColumn("lSelect", typeof(bool)), new DataColumn("cExcleColName", typeof(string)) });
                DtExcelCol.Columns["lSelect"].DefaultValue = true;
                ClsExcelWork.InstanceClsExcelWor.GetExcelColumns(StrExcelFile, StrExcelSheetName, DtExcelCol, "cExcleColName");
                return DtExcelCol;
            }
            catch (Exception)
            {

                throw;
            }
        }
        public DataTable GetExcelData(string StrExcelFile, string StrExcelSheetName)
        {
            try
            {
                DataTable DtExcelCol = new DataTable();

                DtExcelCol = ClsExcelWork.InstanceClsExcelWor.GetExcelData(StrExcelFile, StrExcelSheetName);
                return DtExcelCol;
            }
            catch (Exception)
            {

                throw;
            }
        }
        public void ConvertExcelToText(string StrTextFile, DataView DvExcelData, int iTextSpace)
        {
            try
            {

                
                if (File.Exists(StrTextFile))
                {
                    File.Delete(StrTextFile);
                }
                List<string> LstTextLine = new List<string>();
                Dictionary<string, int> DicColMaxLen = new Dictionary<string, int>();
                foreach (DataColumn DcCol in DvExcelData.ToTable().Columns)
                {

                    int maxStringLength = DvExcelData.ToTable().AsEnumerable()
                                .Select(row => row[DcCol.ColumnName]).Max(str => str.ToString().Length);


                    DicColMaxLen.Add(DcCol.ColumnName, DcCol.ColumnName.Length > maxStringLength ? DcCol.ColumnName.Length : maxStringLength);
                }

                using (StreamWriter sw = File.CreateText(StrTextFile))
                {


                    string StrTextLine = "";
                    foreach (KeyValuePair<string, int> DicColItem in DicColMaxLen)
                    {
                        string StrTextValue = "";
                        int iStringLength = ClsUtility._IClsUtility.ConvertDbNullString(DicColItem.Key).Length;

                        int iMaxStringLength = DicColItem.Value < iStringLength ? iStringLength : DicColItem.Value;
                        int iLengthDiffrance = iMaxStringLength - iStringLength;
                        StrTextValue = ClsUtility._IClsUtility.ConvertDbNullString(DicColItem.Key);
                        StrTextValue = StrTextValue + new String(' ', (iLengthDiffrance + iTextSpace));
                        StrTextLine = string.Concat(StrTextLine, StrTextValue);


                    }
                    sw.WriteLine(StrTextLine);

                    foreach (DataRowView DrvExcleData in DvExcelData)
                    {
                        StrTextLine = "";
                        foreach (KeyValuePair<string, int> DicColItem in DicColMaxLen)
                        {
                            string StrTextValue = "";
                            int iStringLength = ClsUtility._IClsUtility.ConvertDbNullString(DrvExcleData[DicColItem.Key]).Length;
                            int iMaxStringLength = DicColItem.Value;
                            int iLengthDiffrance = iMaxStringLength - iStringLength;
                            StrTextValue = ClsUtility._IClsUtility.ConvertDbNullString(DrvExcleData[DicColItem.Key]);
                            StrTextValue = StrTextValue + new String(' ', (iLengthDiffrance + iTextSpace));
                            StrTextLine = string.Concat(StrTextLine, StrTextValue);


                        }
                        sw.WriteLine(StrTextLine);
                    }

                }

            }
            catch (Exception)
            {

                throw;
            }
        }
        private void CheckFile(string StrFileName)
        {
            try
            {
                if (!File.Exists(StrFileName))
                {
                    File.Create(StrFileName);
                }
                else
                {
                    File.Delete(StrFileName);
                    File.Create(StrFileName);
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
