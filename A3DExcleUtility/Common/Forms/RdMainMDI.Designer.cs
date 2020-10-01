namespace A3DExcleUtility.Common.Forms
{
    partial class RdMainMDI
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RdMainMDI));
            this.RdStatusStrip = new Telerik.WinControls.UI.RadStatusStrip();
            this.RdSplitBtnTheam = new Telerik.WinControls.UI.RadSplitButtonElement();
            this.RdSpltBtnItemThemeFluent = new Telerik.WinControls.UI.RadMenuItem();
            this.RdSpltBtnItemThemeFluentDark = new Telerik.WinControls.UI.RadMenuItem();
            this.commandBarSeparator1 = new Telerik.WinControls.UI.CommandBarSeparator();
            this.RadLblMessage = new Telerik.WinControls.UI.RadLabelElement();
            this.commandBarSeparator2 = new Telerik.WinControls.UI.CommandBarSeparator();
            this.RdProgressBar = new Telerik.WinControls.UI.RadProgressBarElement();
            this.panel1 = new System.Windows.Forms.Panel();
            this.RdDockMain = new Telerik.WinControls.UI.Docking.RadDock();
            this.documentContainer1 = new Telerik.WinControls.UI.Docking.DocumentContainer();
            this.RdRibBarMain = new Telerik.WinControls.UI.RadRibbonBar();
            this.RdRibTabExcelWork = new Telerik.WinControls.UI.RibbonTab();
            this.RdRibBarGrpConverExcle = new Telerik.WinControls.UI.RadRibbonBarGroup();
            this.RdBtnConvertExcleToText = new Telerik.WinControls.UI.RadButtonElement();
            this.RdRibTabAboutUs = new Telerik.WinControls.UI.RibbonTab();
            this.RdRibBarGrpSupport = new Telerik.WinControls.UI.RadRibbonBarGroup();
            this.RdBtnAboutUs = new Telerik.WinControls.UI.RadButtonElement();
            ((System.ComponentModel.ISupportInitialize)(this.RdStatusStrip)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.RdDockMain)).BeginInit();
            this.RdDockMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.documentContainer1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RdRibBarMain)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // RdStatusStrip
            // 
            this.RdStatusStrip.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.RdSplitBtnTheam,
            this.commandBarSeparator1,
            this.RadLblMessage,
            this.commandBarSeparator2,
            this.RdProgressBar});
            this.RdStatusStrip.Location = new System.Drawing.Point(0, 381);
            this.RdStatusStrip.Name = "RdStatusStrip";
            this.RdStatusStrip.Size = new System.Drawing.Size(678, 26);
            this.RdStatusStrip.SizingGrip = false;
            this.RdStatusStrip.TabIndex = 1;
            // 
            // RdSplitBtnTheam
            // 
            this.RdSplitBtnTheam.ArrowButtonMinSize = new System.Drawing.Size(12, 12);
            this.RdSplitBtnTheam.DefaultItem = null;
            this.RdSplitBtnTheam.DropDownDirection = Telerik.WinControls.UI.RadDirection.Down;
            this.RdSplitBtnTheam.ExpandArrowButton = false;
            this.RdSplitBtnTheam.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.RdSpltBtnItemThemeFluent,
            this.RdSpltBtnItemThemeFluentDark});
            this.RdSplitBtnTheam.Name = "RdSplitBtnTheam";
            this.RdStatusStrip.SetSpring(this.RdSplitBtnTheam, false);
            this.RdSplitBtnTheam.Text = "";
            this.RdSplitBtnTheam.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            // 
            // RdSpltBtnItemThemeFluent
            // 
            this.RdSpltBtnItemThemeFluent.Image = global::A3DExcleUtility.Properties.Resources.LightTheam24X24;
            this.RdSpltBtnItemThemeFluent.Name = "RdSpltBtnItemThemeFluent";
            this.RdSpltBtnItemThemeFluent.Tag = "Fluent";
            this.RdSpltBtnItemThemeFluent.Text = "   Light";
            this.RdSpltBtnItemThemeFluent.Click += new System.EventHandler(this.RdSpltBtnItemThemeFluent_Click);
            // 
            // RdSpltBtnItemThemeFluentDark
            // 
            this.RdSpltBtnItemThemeFluentDark.Image = global::A3DExcleUtility.Properties.Resources.DarkTheam24X24;
            this.RdSpltBtnItemThemeFluentDark.Name = "RdSpltBtnItemThemeFluentDark";
            this.RdSpltBtnItemThemeFluentDark.Tag = "FluentDark";
            this.RdSpltBtnItemThemeFluentDark.Text = "   Dark";
            this.RdSpltBtnItemThemeFluentDark.Click += new System.EventHandler(this.RdSpltBtnItemThemeFluent_Click);
            // 
            // commandBarSeparator1
            // 
            this.commandBarSeparator1.Name = "commandBarSeparator1";
            this.RdStatusStrip.SetSpring(this.commandBarSeparator1, false);
            this.commandBarSeparator1.VisibleInOverflowMenu = false;
            // 
            // RadLblMessage
            // 
            this.RadLblMessage.Name = "RadLblMessage";
            this.RdStatusStrip.SetSpring(this.RadLblMessage, false);
            this.RadLblMessage.Text = "..";
            this.RadLblMessage.TextWrap = true;
            // 
            // commandBarSeparator2
            // 
            this.commandBarSeparator2.Name = "commandBarSeparator2";
            this.RdStatusStrip.SetSpring(this.commandBarSeparator2, false);
            this.commandBarSeparator2.VisibleInOverflowMenu = false;
            // 
            // RdProgressBar
            // 
            this.RdProgressBar.Name = "RdProgressBar";
            this.RdProgressBar.SeparatorColor1 = System.Drawing.Color.White;
            this.RdProgressBar.SeparatorColor2 = System.Drawing.Color.White;
            this.RdProgressBar.SeparatorColor3 = System.Drawing.Color.White;
            this.RdProgressBar.SeparatorColor4 = System.Drawing.Color.White;
            this.RdProgressBar.SeparatorGradientAngle = 0;
            this.RdProgressBar.SeparatorGradientPercentage1 = 0.4F;
            this.RdProgressBar.SeparatorGradientPercentage2 = 0.6F;
            this.RdProgressBar.SeparatorNumberOfColors = 2;
            this.RdStatusStrip.SetSpring(this.RdProgressBar, true);
            this.RdProgressBar.StepWidth = 14;
            this.RdProgressBar.SweepAngle = 90;
            this.RdProgressBar.Text = "";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.RdDockMain);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 173);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(678, 208);
            this.panel1.TabIndex = 2;
            // 
            // RdDockMain
            // 
            this.RdDockMain.Controls.Add(this.documentContainer1);
            this.RdDockMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.RdDockMain.IsCleanUpTarget = true;
            this.RdDockMain.Location = new System.Drawing.Point(0, 0);
            this.RdDockMain.MainDocumentContainer = this.documentContainer1;
            this.RdDockMain.Name = "RdDockMain";
            // 
            // 
            // 
            this.RdDockMain.RootElement.MinSize = new System.Drawing.Size(25, 25);
            this.RdDockMain.Size = new System.Drawing.Size(678, 208);
            this.RdDockMain.TabIndex = 0;
            this.RdDockMain.TabStop = false;
            this.RdDockMain.DockWindowAdded += new Telerik.WinControls.UI.Docking.DockWindowEventHandler(this.RdMainDock_DockWindowAdded);
            // 
            // documentContainer1
            // 
            this.documentContainer1.Name = "documentContainer1";
            // 
            // 
            // 
            this.documentContainer1.RootElement.MinSize = new System.Drawing.Size(25, 25);
            this.documentContainer1.SizeInfo.SizeMode = Telerik.WinControls.UI.Docking.SplitPanelSizeMode.Fill;
            // 
            // RdRibBarMain
            // 
            this.RdRibBarMain.CommandTabs.AddRange(new Telerik.WinControls.RadItem[] {
            this.RdRibTabExcelWork,
            this.RdRibTabAboutUs});
            // 
            // 
            // 
            this.RdRibBarMain.ExitButton.Text = "Exit";
            this.RdRibBarMain.ExitButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.RdRibBarMain.LocalizationSettings.LayoutModeText = "Simplified Layout";
            this.RdRibBarMain.Location = new System.Drawing.Point(0, 0);
            this.RdRibBarMain.Name = "RdRibBarMain";
            // 
            // 
            // 
            this.RdRibBarMain.OptionsButton.Text = "Options";
            this.RdRibBarMain.OptionsButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            // 
            // 
            // 
            this.RdRibBarMain.RootElement.AutoSizeMode = Telerik.WinControls.RadAutoSizeMode.WrapAroundChildren;
            this.RdRibBarMain.Size = new System.Drawing.Size(678, 173);
            this.RdRibBarMain.StartButtonImage = global::A3DExcleUtility.Properties.Resources.A3D24X24;
            this.RdRibBarMain.TabIndex = 0;
            this.RdRibBarMain.Click += new System.EventHandler(this.RdRibBarMain_Click);
            ((Telerik.WinControls.UI.RadRibbonBarElement)(this.RdRibBarMain.GetChildAt(0))).Text = "";
            ((Telerik.WinControls.UI.RadQuickAccessToolBar)(this.RdRibBarMain.GetChildAt(0).GetChildAt(2))).Image = null;
            ((Telerik.WinControls.UI.RadRibbonBarCaption)(this.RdRibBarMain.GetChildAt(0).GetChildAt(3))).DrawText = false;
            // 
            // RdRibTabExcelWork
            // 
            this.RdRibTabExcelWork.Image = global::A3DExcleUtility.Properties.Resources.Excel24X24;
            this.RdRibTabExcelWork.IsSelected = true;
            this.RdRibTabExcelWork.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.RdRibBarGrpConverExcle});
            this.RdRibTabExcelWork.Name = "RdRibTabExcelWork";
            this.RdRibTabExcelWork.Text = "Excel Work";
            this.RdRibTabExcelWork.UseMnemonic = false;
            // 
            // RdRibBarGrpConverExcle
            // 
            this.RdRibBarGrpConverExcle.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.RdBtnConvertExcleToText});
            this.RdRibBarGrpConverExcle.Name = "RdRibBarGrpConverExcle";
            this.RdRibBarGrpConverExcle.Text = "Convert Excel";
            // 
            // RdBtnConvertExcleToText
            // 
            this.RdBtnConvertExcleToText.Image = global::A3DExcleUtility.Properties.Resources.ConvertFile48X48;
            this.RdBtnConvertExcleToText.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            this.RdBtnConvertExcleToText.Name = "RdBtnConvertExcleToText";
            this.RdBtnConvertExcleToText.Tag = "RdFrmExcelToText";
            this.RdBtnConvertExcleToText.Text = "Excel To Text";
            this.RdBtnConvertExcleToText.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.RdBtnConvertExcleToText.Click += new System.EventHandler(this.RdBtnConvertExcleToText_Click);
            // 
            // RdRibTabAboutUs
            // 
            this.RdRibTabAboutUs.Image = global::A3DExcleUtility.Properties.Resources.info;
            this.RdRibTabAboutUs.IsSelected = false;
            this.RdRibTabAboutUs.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.RdRibBarGrpSupport});
            this.RdRibTabAboutUs.Name = "RdRibTabAboutUs";
            this.RdRibTabAboutUs.Text = "About Us";
            this.RdRibTabAboutUs.UseMnemonic = false;
            // 
            // RdRibBarGrpSupport
            // 
            this.RdRibBarGrpSupport.ClipDrawing = true;
            this.RdRibBarGrpSupport.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.RdBtnAboutUs});
            this.RdRibBarGrpSupport.Name = "RdRibBarGrpSupport";
            this.RdRibBarGrpSupport.Text = "About Us";
            // 
            // RdBtnAboutUs
            // 
            this.RdBtnAboutUs.Image = global::A3DExcleUtility.Properties.Resources.AboutUs48X48;
            this.RdBtnAboutUs.Name = "RdBtnAboutUs";
            this.RdBtnAboutUs.Text = "Support";
            this.RdBtnAboutUs.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.RdBtnAboutUs.Click += new System.EventHandler(this.RdBtnAboutUs_Click);
            // 
            // RdMainMDI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(678, 407);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.RdStatusStrip);
            this.Controls.Add(this.RdRibBarMain);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.IsMdiContainer = true;
            this.MainMenuStrip = null;
            this.Name = "RdMainMDI";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.RdStatusStrip)).EndInit();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.RdDockMain)).EndInit();
            this.RdDockMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.documentContainer1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RdRibBarMain)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Telerik.WinControls.UI.RadRibbonBar RdRibBarMain;
        private Telerik.WinControls.UI.RadStatusStrip RdStatusStrip;
        private System.Windows.Forms.Panel panel1;
        private Telerik.WinControls.UI.RibbonTab RdRibTabExcelWork;
        private Telerik.WinControls.UI.RibbonTab RdRibTabAboutUs;
        private Telerik.WinControls.UI.RadRibbonBarGroup RdRibBarGrpConverExcle;
        private Telerik.WinControls.UI.RadButtonElement RdBtnConvertExcleToText;
        private Telerik.WinControls.UI.RadRibbonBarGroup RdRibBarGrpSupport;
        private Telerik.WinControls.UI.RadButtonElement RdBtnAboutUs;
        private Telerik.WinControls.UI.Docking.RadDock RdDockMain;
        private Telerik.WinControls.UI.Docking.DocumentContainer documentContainer1;
        private Telerik.WinControls.UI.RadSplitButtonElement RdSplitBtnTheam;
        private Telerik.WinControls.UI.RadMenuItem RdSpltBtnItemThemeFluent;
        private Telerik.WinControls.UI.RadMenuItem RdSpltBtnItemThemeFluentDark;
        private Telerik.WinControls.UI.CommandBarSeparator commandBarSeparator1;
        private Telerik.WinControls.UI.RadLabelElement RadLblMessage;
        private Telerik.WinControls.UI.CommandBarSeparator commandBarSeparator2;
        private Telerik.WinControls.UI.RadProgressBarElement RdProgressBar;
    }
}
