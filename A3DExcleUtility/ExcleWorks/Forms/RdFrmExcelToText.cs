using A3DExcleUtility.ExcleWorks.Classes;
using A3DWinUtility;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using Telerik.WinControls.UI;

namespace A3DExcleUtility.ExcleWorks.Forms
{
    public partial class RdFrmExcelToText : Telerik.WinControls.UI.RadForm
    {
        DataTable DtExcelColumns = new DataTable();
        DataTable DtExcelData = new DataTable();
        public RdFrmExcelToText()
        {
            InitializeComponent();
            RdSpltContainerMain.UseSplitterButtons = true;
            //RdSpltContainerMain.EnableCollapsing = true;
            RdTxtExcelSpace.TextBoxElement.ToolTipText = "Select Excel File";
            RdTxtExcelSpace.TextBoxElement.ToolTipText = "Space Between Excel Columns Default is 4 Character";
            RdDdExcelSheet.DropDownListElement.ToolTipText = "Select Excel Sheet";
            RdBtnConvert.ButtonElement.ToolTipText = "Convert Excel File To Text";
            RdBtnSelectFile.ButtonElement.ToolTipText = "Select Excel File";
            RdBtnSelectAll.ButtonElement.ToolTipText = "Select All Excel Columns";
            RdBtnUnSelectAll.ButtonElement.ToolTipText = "Un-Select All Excel Columns";
        }

        private void RdFrmExcelToText_Load(object sender, EventArgs e)
        {

        }

        private void RdBtnSelectFile_Click(object sender, EventArgs e)
        {
            try
            {
                using (RadOpenFileDialog OpenDlg = new RadOpenFileDialog())
                {

                    DtExcelColumns = new DataTable();
                    DtExcelData = new DataTable();
                    RdGrdExcelCol.DataSource = null;
                    RdGrdExcelData.DataSource = null;
                    RdGrdExcelCol.Rows.Clear();
                    RdGrdExcelData.Rows.Clear();
                    RdTxtSelectFile.Text = string.Empty;
                    RdDdExcelSheet.Items.Clear();
                    RdDdExcelSheet.Text = string.Empty;
                    RdTxtExcelSpace.Text = "4";

                    OpenDlg.DefaultExt = ".xlsx";
                    OpenDlg.Filter = "Excel File(*.xlsx)|*.xlsx";
                    OpenDlg.MultiSelect = false;
                    OpenDlg.InitialDirectory = Path.GetFullPath(Environment.SpecialFolder.Desktop.ToString());
                    if (OpenDlg.ShowDialog() == DialogResult.OK)
                    {
                        Cursor = Cursors.WaitCursor;

                        RdTxtSelectFile.Text = OpenDlg.FileName;
                        GetExcelSheet(RdTxtSelectFile.Text.Trim());
                        Cursor = Cursors.Default;
                    }
                }
            }
            catch (Exception ex)
            {
                ClsMessage._IClsMessage.ProjectExceptionMessage(ex);
                Cursor = Cursors.Default;
            }
        }
        private void GetExcelSheet(string StrExcelFile)
        {
            try
            {
                var lstsheet = ClsExcelToText._IClsExcelToText.GetExcelSheet(StrExcelFile);
                RdDdExcelSheet.Items.Clear();
                foreach (var item in lstsheet)
                {
                    RdDdExcelSheet.Items.Add(item);
                }
                RdDdExcelSheet.SelectedIndex = 0;
            }
            catch (Exception)
            {

                throw;
            }
        }
        private void GetExcelColumns()
        {
            try
            {
                if (!string.IsNullOrEmpty(RdTxtSelectFile.Text) && !string.IsNullOrEmpty(RdDdExcelSheet.Text))
                {


                    DtExcelColumns = ClsExcelToText._IClsExcelToText.GetExcelColumns(RdTxtSelectFile.Text.Trim(), RdDdExcelSheet.Text.Trim());
                    RdGrdExcelCol.DataSource = DtExcelColumns.DefaultView;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        private void GetExcelData()
        {
            try
            {
                if (!string.IsNullOrEmpty(RdTxtSelectFile.Text) && !string.IsNullOrEmpty(RdDdExcelSheet.Text))
                {

                    DtExcelData = ClsExcelToText._IClsExcelToText.GetExcelData(RdTxtSelectFile.Text.Trim(), RdDdExcelSheet.Text.Trim());
                    RdGrdExcelData.DataSource = DtExcelData.DefaultView;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        private void RdDdExcelSheet_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
        {
            try
            {
                if (RdDdExcelSheet.Text != null && RdDdExcelSheet.Text != "")
                {
                    Cursor = Cursors.WaitCursor;
                    GetExcelColumns();
                    GetExcelData();
                    Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {
                ClsMessage._IClsMessage.ProjectExceptionMessage(ex);
                Cursor = Cursors.Default;
            }
        }

        private void RdBtnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                var DcSelectedCol = DtExcelColumns.AsEnumerable().Where(xs => xs.Field<bool>("lSelect") == true).Select(xl => xl["cExcleColName"].ToString()).ToList();
                if (DcSelectedCol != null && DcSelectedCol.Count > 0)
                {
                    DataView DvSelectedData = new DataView();
                    DataSet ds = new DataSet();
                    DvSelectedData = ds.DefaultViewManager.CreateDataView(DtExcelData);
                    DvSelectedData = DvSelectedData.ToTable(true, DcSelectedCol.ToArray()).DefaultView;
                    using (RadSaveFileDialog SavDlg = new RadSaveFileDialog())
                    {
                        SavDlg.Filter = "Text File(*.txt)|*.txt";
                        SavDlg.DefaultExt = ".txt";
                        SavDlg.RestoreDirectory = true;
                        SavDlg.InitialDirectory = Path.GetFullPath(Environment.SpecialFolder.Desktop.ToString());
                        SavDlg.FileName = RdDdExcelSheet.Text;
                        if (SavDlg.ShowDialog() == DialogResult.OK)
                        {
                            ClsExcelToText._IClsExcelToText.ConvertExcelToText(SavDlg.FileName, DvSelectedData, RdTxtExcelSpace.Text != null && RdTxtExcelSpace.Text.Trim() != "" ? Convert.ToInt32(RdTxtExcelSpace.Text.Trim()) : 4);
                        }
                        if (ClsMessage._IClsMessage.showQuestionMessage("Text File Create!!" + Environment.NewLine + "Do You Want To Open It?") == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(SavDlg.FileName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                ClsMessage._IClsMessage.ProjectExceptionMessage(ex);
            }
        }
    }
}
