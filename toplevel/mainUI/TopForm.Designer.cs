namespace mainUI
{
    partial class TopForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TopForm));
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.contextMenuStripMain = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.executeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pauseToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stopToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.selectAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.selectFromHereToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolTipMain = new System.Windows.Forms.ToolTip(this.components);
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.ribbonMain = new System.Windows.Forms.Ribbon();
            this.ribbonOrbMenuItemOpen = new System.Windows.Forms.RibbonOrbMenuItem();
            this.ribbonOrbMenuItemSave = new System.Windows.Forms.RibbonButton();
            this.ribbonOrbMenuItemSaveAs = new System.Windows.Forms.RibbonOrbMenuItem();
            this.ribbonOrbMenuItemClose = new System.Windows.Forms.RibbonOrbMenuItem();
            this.ribbonOrbMenuItemExit = new System.Windows.Forms.RibbonOrbMenuItem();
            this.ribbonOrbOptionButtonAbout = new System.Windows.Forms.RibbonOrbOptionButton();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.contextMenuStripMain.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView
            // 
            this.dataGridView.ContextMenuStrip = this.contextMenuStripMain;
            this.dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView.Location = new System.Drawing.Point(0, 141);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView.Size = new System.Drawing.Size(960, 374);
            this.dataGridView.TabIndex = 0;
            this.dataGridView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellClick);
            this.dataGridView.CellMouseUp += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView_CellMouseUp);
            // 
            // contextMenuStripMain
            // 
            this.contextMenuStripMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.executeToolStripMenuItem,
            this.pauseToolStripMenuItem,
            this.stopToolStripMenuItem,
            this.selectAllToolStripMenuItem,
            this.selectFromHereToolStripMenuItem,
            this.deleteRowToolStripMenuItem});
            this.contextMenuStripMain.Name = "contextMenuStrip1";
            this.contextMenuStripMain.Size = new System.Drawing.Size(165, 136);
            // 
            // executeToolStripMenuItem
            // 
            this.executeToolStripMenuItem.Name = "executeToolStripMenuItem";
            this.executeToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
            this.executeToolStripMenuItem.Text = "Execute";
            this.executeToolStripMenuItem.Click += new System.EventHandler(this.executeToolStripMenuItem_Click);
            // 
            // pauseToolStripMenuItem
            // 
            this.pauseToolStripMenuItem.Name = "pauseToolStripMenuItem";
            this.pauseToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
            this.pauseToolStripMenuItem.Text = "Pause";
            this.pauseToolStripMenuItem.Click += new System.EventHandler(this.pauseToolStripMenuItem_Click);
            // 
            // stopToolStripMenuItem
            // 
            this.stopToolStripMenuItem.Name = "stopToolStripMenuItem";
            this.stopToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
            this.stopToolStripMenuItem.Text = "Stop";
            this.stopToolStripMenuItem.Click += new System.EventHandler(this.stopToolStripMenuItem_Click);
            // 
            // selectAllToolStripMenuItem
            // 
            this.selectAllToolStripMenuItem.Name = "selectAllToolStripMenuItem";
            this.selectAllToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
            this.selectAllToolStripMenuItem.Text = "Select All";
            this.selectAllToolStripMenuItem.Click += new System.EventHandler(this.selectAllToolStripMenuItem_Click);
            // 
            // selectFromHereToolStripMenuItem
            // 
            this.selectFromHereToolStripMenuItem.Name = "selectFromHereToolStripMenuItem";
            this.selectFromHereToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
            this.selectFromHereToolStripMenuItem.Text = "Select From Here";
            this.selectFromHereToolStripMenuItem.Click += new System.EventHandler(this.selectFromHereToolStripMenuItem_Click);
            // 
            // deleteRowToolStripMenuItem
            // 
            this.deleteRowToolStripMenuItem.Name = "deleteRowToolStripMenuItem";
            this.deleteRowToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
            this.deleteRowToolStripMenuItem.Text = "Delete Row";
            this.deleteRowToolStripMenuItem.Click += new System.EventHandler(this.deleteRowToolStripMenuItem_Click);
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
            this.openFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 515);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.ShowItemToolTips = true;
            this.statusStrip1.Size = new System.Drawing.Size(960, 22);
            this.statusStrip1.TabIndex = 2;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.AutoToolTip = true;
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel.Text = "toolStripStatusLabel1";
            // 
            // ribbonMain
            // 
            this.ribbonMain.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.ribbonMain.Location = new System.Drawing.Point(0, 0);
            this.ribbonMain.Minimized = false;
            this.ribbonMain.Name = "ribbonMain";
            // 
            // 
            // 
            this.ribbonMain.OrbDropDown.BorderRoundness = 8;
            this.ribbonMain.OrbDropDown.Location = new System.Drawing.Point(0, 0);
            this.ribbonMain.OrbDropDown.MenuItems.Add(this.ribbonOrbMenuItemOpen);
            this.ribbonMain.OrbDropDown.MenuItems.Add(this.ribbonOrbMenuItemSave);
            this.ribbonMain.OrbDropDown.MenuItems.Add(this.ribbonOrbMenuItemSaveAs);
            this.ribbonMain.OrbDropDown.MenuItems.Add(this.ribbonOrbMenuItemClose);
            this.ribbonMain.OrbDropDown.MenuItems.Add(this.ribbonOrbMenuItemExit);
            this.ribbonMain.OrbDropDown.Name = "";
            this.ribbonMain.OrbDropDown.OptionItems.Add(this.ribbonOrbOptionButtonAbout);
            this.ribbonMain.OrbDropDown.RecentItemsCaption = "Recent Testcases Loaded";
            this.ribbonMain.OrbDropDown.Size = new System.Drawing.Size(527, 351);
            this.ribbonMain.OrbDropDown.TabIndex = 0;
            this.ribbonMain.OrbImage = null;
            this.ribbonMain.OrbStyle = System.Windows.Forms.RibbonOrbStyle.Office_2013;
            this.ribbonMain.OrbText = "File";
            // 
            // 
            // 
            this.ribbonMain.QuickAcessToolbar.DropDownButtonVisible = false;
            this.ribbonMain.RibbonTabFont = new System.Drawing.Font("Trebuchet MS", 9F);
            this.ribbonMain.Size = new System.Drawing.Size(960, 141);
            this.ribbonMain.TabIndex = 1;
            this.ribbonMain.TabsMargin = new System.Windows.Forms.Padding(12, 26, 20, 0);
            this.ribbonMain.Text = "ribbonMain";
            this.ribbonMain.ThemeColor = System.Windows.Forms.RibbonTheme.Blue;
            // 
            // ribbonOrbMenuItemOpen
            // 
            this.ribbonOrbMenuItemOpen.DropDownArrowDirection = System.Windows.Forms.RibbonArrowDirection.Left;
            this.ribbonOrbMenuItemOpen.Image = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItemOpen.Image")));
            this.ribbonOrbMenuItemOpen.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItemOpen.SmallImage")));
            this.ribbonOrbMenuItemOpen.Text = "Open";
            this.ribbonOrbMenuItemOpen.Click += new System.EventHandler(this.ribbonOrbMenuItemOpen_Click);
            // 
            // ribbonOrbMenuItemSave
            // 
            this.ribbonOrbMenuItemSave.Image = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItemSave.Image")));
            this.ribbonOrbMenuItemSave.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItemSave.SmallImage")));
            this.ribbonOrbMenuItemSave.Text = "Save";
            this.ribbonOrbMenuItemSave.Click += new System.EventHandler(this.ribbonOrbMenuItemSave_Click);
            // 
            // ribbonOrbMenuItemSaveAs
            // 
            this.ribbonOrbMenuItemSaveAs.DropDownArrowDirection = System.Windows.Forms.RibbonArrowDirection.Left;
            this.ribbonOrbMenuItemSaveAs.Image = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItemSaveAs.Image")));
            this.ribbonOrbMenuItemSaveAs.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItemSaveAs.SmallImage")));
            this.ribbonOrbMenuItemSaveAs.Text = "Save As";
            this.ribbonOrbMenuItemSaveAs.Click += new System.EventHandler(this.ribbonOrbMenuItemSaveAs_Click);
            // 
            // ribbonOrbMenuItemClose
            // 
            this.ribbonOrbMenuItemClose.DropDownArrowDirection = System.Windows.Forms.RibbonArrowDirection.Left;
            this.ribbonOrbMenuItemClose.Image = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItemClose.Image")));
            this.ribbonOrbMenuItemClose.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItemClose.SmallImage")));
            this.ribbonOrbMenuItemClose.Text = "Close";
            this.ribbonOrbMenuItemClose.Click += new System.EventHandler(this.ribbonOrbMenuItemClose_Click);
            // 
            // ribbonOrbMenuItemExit
            // 
            this.ribbonOrbMenuItemExit.DropDownArrowDirection = System.Windows.Forms.RibbonArrowDirection.Left;
            this.ribbonOrbMenuItemExit.Image = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItemExit.Image")));
            this.ribbonOrbMenuItemExit.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonOrbMenuItemExit.SmallImage")));
            this.ribbonOrbMenuItemExit.Text = "Exit";
            this.ribbonOrbMenuItemExit.Click += new System.EventHandler(this.ribbonOrbMenuItemExit_Click);
            // 
            // ribbonOrbOptionButtonAbout
            // 
            this.ribbonOrbOptionButtonAbout.Image = ((System.Drawing.Image)(resources.GetObject("ribbonOrbOptionButtonAbout.Image")));
            this.ribbonOrbOptionButtonAbout.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonOrbOptionButtonAbout.SmallImage")));
            this.ribbonOrbOptionButtonAbout.Text = "About";
            this.ribbonOrbOptionButtonAbout.Click += new System.EventHandler(this.ribbonOrbOptionButtonAbout_Click);
            // 
            // TopForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(960, 537);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.ribbonMain);
            this.Controls.Add(this.statusStrip1);
            this.Name = "TopForm";
            this.Load += new System.EventHandler(this.TopForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.contextMenuStripMain.ResumeLayout(false);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        // create on the fly
        private System.Windows.Forms.RibbonTab ribbonTab2;
        private System.Windows.Forms.RibbonPanel ribbonPanel2;

        private System.Windows.Forms.DataGridView dataGridView;

        private System.Windows.Forms.Ribbon ribbonMain;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripMain;
        private System.Windows.Forms.ToolStripMenuItem executeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pauseToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem stopToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem selectAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem selectFromHereToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deleteRowToolStripMenuItem;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.ToolTip toolTipMain;
        private System.Windows.Forms.RibbonOrbMenuItem ribbonOrbMenuItemOpen;
        private System.Windows.Forms.RibbonOrbMenuItem ribbonOrbMenuItemSaveAs;
        private System.Windows.Forms.RibbonOrbMenuItem ribbonOrbMenuItemClose;
        private System.Windows.Forms.RibbonOrbMenuItem ribbonOrbMenuItemExit;
        private System.Windows.Forms.RibbonOrbOptionButton ribbonOrbOptionButtonAbout;
        private System.Windows.Forms.RibbonButton ribbonOrbMenuItemSave;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
    }
}

