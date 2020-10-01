using A3DWinUtility;
using FontAwesome.Sharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using Telerik.WinControls.UI;
using Telerik.WinControls.UI.Docking;

namespace A3DExcleUtility.Common.Forms
{
    public partial class RdMainMDI : RadRibbonForm
    {
        private Random rnd = new Random();
        public string cFormIcon { get; set; } = "Circle";
        public RdMainMDI()
        {
            InitializeComponent();
            Text = "A3D Excel Utility";
            ShowIcon = true;
            Icon = Properties.Resources.A3D;
            AllowAero = false;
            RdRibBarMain.RibbonBarElement.ApplicationButtonElement.Visibility = Telerik.WinControls.ElementVisibility.Hidden;
            RdRibBarMain.RibbonBarElement.IconPrimitive.Visibility = ElementVisibility.Visible;
            RdDockMain.MdiChildrenDockType = DockType.Document;
            RdDockMain.AutoDetectMdiChildren = true;

            SetSplitBtn(RdSpltBtnItemThemeFluent);
        }

        private void RdRibBarMain_Click(object sender, EventArgs e)
        {

        }

        private void RdBtnAboutUs_Click(object sender, EventArgs e)
        {
            RdFrmAboutUs rdFrmAboutUs = new RdFrmAboutUs();
            rdFrmAboutUs.ShowDialog(this);

        }
        public RadForm TryGetFormByName(string frmname)
        {
            var formType = Assembly.GetExecutingAssembly().GetTypes().FirstOrDefault(a => (a.BaseType == typeof(RadForm) || a.BaseType == typeof(Form)) && a.Name == frmname);

            if (formType == null) // If there is no form with the given frmname
                return null;


            return (RadForm)Activator.CreateInstance(formType);

        }

        private void OpenForm(string strMenuName, string strFormName)
        {
            RadForm _RadForm = new RadForm();
            _RadForm = TryGetFormByName(strFormName);
            HostWindow _HostWindowForm = null;
            if (ClsUtility._IClsUtility.IsFormOpen(_RadForm.GetType(), RdDockMain, out _HostWindowForm))
            {
                _HostWindowForm.ActivateAsMdiChild();
                _RadForm.Activate();
                RdDockMain.ActivateWindow(_HostWindowForm);
                RdDockMain.ActivateMdiChild(_RadForm);

            }
            else
            {
                //var icon = IconChar.FileExcel;
                Color randomColor = Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));
                cFormIcon = _RadForm.Tag != null && _RadForm.Tag.ToString().Trim() != "" ? _RadForm.Tag.ToString().Trim() : "Circle";
                var iconImage = ((IconChar)Enum.Parse(typeof(IconChar), cFormIcon)).ToBitmap(16, randomColor);
                _RadForm.MdiParent = this;
                _RadForm.Show();
                _RadForm.Activate();
                //_RadForm.Icon = iconImage;
                RdDockMain.ActivateMdiChild(_RadForm);
                foreach (DockWindow dw in RdDockMain.DockWindows)
                {
                    dw.AllowedDockState = ~AllowedDockState.Floating;
                }
            }
        }
        private void RdMainDock_DockWindowAdded(object sender, DockWindowEventArgs e)

        {
            try
            {

                //RadForm form = ((RadForm)RdMainDock.ActiveWindow.ActiveControl);
                //Image img = new Bitmap(form.Icon.ToBitmap(), new Size(16, 16));
                e.DockWindow.TabStripItem.Image = ((IconChar)Enum.Parse(typeof(IconChar), cFormIcon)).ToBitmap(16, Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256))); ;
            }
            catch (Exception ex)
            {

                throw;
            }
            //RadForm form = ((RadForm)radDock1.ActiveWindow.ActiveControl);
            //Image img = new Bitmap(form.Icon.ToBitmap(), new Size(16, 16));
            //e.DockWindow.TabStripItem.Image = img;
        }
        private void RdBtnConvertExcleToText_Click(object sender, EventArgs e)
        {
            try
            {
                RadButtonElement _RadButtonElement = (RadButtonElement)sender;
                if ((!string.IsNullOrEmpty(_RadButtonElement.Tag.ToString()) && _RadButtonElement.Tag.ToString() != ""))
                {
                    OpenForm(_RadButtonElement.Text, _RadButtonElement.Tag.ToString());
                }
                if ((!string.IsNullOrEmpty(_RadButtonElement.Text) && _RadButtonElement.Text != "") && (_RadButtonElement.Text.ToUpper() == "LOG OUT & EXIST"))
                {

                    DialogResult dgRsult = MessageBox.Show(this, "Do you want exit or log out?" + Environment.NewLine + "Click Yes for exit No for Log Out.", "SBS ATM", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3);
                    switch (dgRsult)
                    {
                        case DialogResult.Yes:
                            Application.Exit();
                            break;
                        case DialogResult.No:
                            Application.Restart();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {


                ClsMessage._IClsMessage.ProjectExceptionMessage(ex.Message);
            }
        }
        private void SetSplitBtn(RadMenuItem radMenuItem)
        {
            try
            {
                RdSplitBtnTheam.DefaultItem = radMenuItem;
                RdSplitBtnTheam.Text = radMenuItem.Text;
                RdSplitBtnTheam.Image = radMenuItem.Image;
            }
            catch (Exception)
            {

                throw;
            }
        }
        private void RdSpltBtnItemThemeFluent_Click(object sender, EventArgs e)
        {
            RadMenuItem radMenuItem = sender as RadMenuItem;
            SetSplitBtn(radMenuItem);
            string TheamName = radMenuItem.Tag != null && radMenuItem.Tag.ToString() != "" ? radMenuItem.Tag.ToString().Trim() : "Fluent";
            switch (TheamName)
            {
                case "FluentDark":
                    ThemeResolutionService.ApplicationThemeName = TheamName;
                    break;
                default:
                    ThemeResolutionService.ApplicationThemeName = TheamName;
                    break;
            }


        }
    }
}
