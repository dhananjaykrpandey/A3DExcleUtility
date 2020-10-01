using A3DExcleUtility.Common.Forms;
using A3DWinUtility;
using System;
using System.Linq;
using System.Windows.Forms;
using Telerik.WinControls;

namespace A3DExcleUtility
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new RadForm1());
            try
            {
                Telerik.WinControls.Themes.FluentTheme FluentTheme = new Telerik.WinControls.Themes.FluentTheme();
                Telerik.WinControls.Themes.FluentDarkTheme fluentDarkTheme = new Telerik.WinControls.Themes.FluentDarkTheme();
                ThemeResolutionService.ApplicationThemeName = "Fluent";
                ThemeResolutionService.AllowAnimations = true;
                //RadMessageBox.SetThemeName(FluentTheme.ThemeName);
                if (ClsUtility._IClsUtility.CheckLicense() == false)
                {
                    A3DLicense.FrmLicense ObjLic = new A3DLicense.FrmLicense();
                    ObjLic.StartPosition = FormStartPosition.CenterScreen;
                    ObjLic.ShowDialog();
                }
                else
                {
                    //InfSQLServices._IInfSQLServices.InfSQLConnectionString = GClsProjectProperties._IGClsProjectProperties.CProjectSqlConnection;
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new RdMainMDI());
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}