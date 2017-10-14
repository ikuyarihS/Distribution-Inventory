using System;

namespace ChiaHang
{
    partial class MainForm
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.navigateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.readPOToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.readFCToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.readFCPlanningToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.readOpenConfigToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.readDeliToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mastahToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fiteMoiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.kgToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.unitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.noSupToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.compactMastahToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cap100ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.seperateFarmsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.seperateThuMuaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.seperateAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.seperateAllOnlyFarmToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.noCapToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupAllNoCapToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.seperateFarmNoCapToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.seperateThuMuaNoCapToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.seperateAllNoCapToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.seperateAllOnlyFarmToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.printToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.purchaseOrderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.horizontallyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.unitToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.kgToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.verticallyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.unitToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.kgToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.compactToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.compactPOunitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.compactPOkgToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnPrintPOReport = new System.Windows.Forms.ToolStripMenuItem();
            this.testStuffToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.testExportLargeExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.reportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.m1ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.printToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mongoDbToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DateFromPicker = new System.Windows.Forms.DateTimePicker();
            this.DateToPicker = new System.Windows.Forms.DateTimePicker();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.progressBarLabel = new System.Windows.Forms.Label();
            this.upperCapBox = new System.Windows.Forms.TextBox();
            this.checkBoxFruitOnly = new System.Windows.Forms.CheckBox();
            this.checkBoxNoFruit = new System.Windows.Forms.CheckBox();
            this.YesNoSubRegionchkBox = new System.Windows.Forms.CheckBox();
            this.richTextBoxOutput = new System.Windows.Forms.RichTextBox();
            this.backgroundCoord = new System.ComponentModel.BackgroundWorker();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(28, 28);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.navigateToolStripMenuItem,
            this.mastahToolStripMenuItem,
            this.printToolStripMenuItem,
            this.testStuffToolStripMenuItem,
            this.reportToolStripMenuItem,
            this.deleteToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1184, 38);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // navigateToolStripMenuItem
            // 
            this.navigateToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.readPOToolStripMenuItem,
            this.readFCToolStripMenuItem,
            this.readFCPlanningToolStripMenuItem,
            this.readOpenConfigToolStripMenuItem,
            this.readDeliToolStripMenuItem});
            this.navigateToolStripMenuItem.Name = "navigateToolStripMenuItem";
            this.navigateToolStripMenuItem.Size = new System.Drawing.Size(185, 34);
            this.navigateToolStripMenuItem.Text = "Update Database";
            // 
            // readPOToolStripMenuItem
            // 
            this.readPOToolStripMenuItem.Name = "readPOToolStripMenuItem";
            this.readPOToolStripMenuItem.Size = new System.Drawing.Size(268, 34);
            this.readPOToolStripMenuItem.Text = "Read PO";
            this.readPOToolStripMenuItem.Click += new System.EventHandler(this.readPOToolStripMenuItem_Click);
            // 
            // readFCToolStripMenuItem
            // 
            this.readFCToolStripMenuItem.Name = "readFCToolStripMenuItem";
            this.readFCToolStripMenuItem.Size = new System.Drawing.Size(268, 34);
            this.readFCToolStripMenuItem.Text = "Read FC";
            this.readFCToolStripMenuItem.Click += new System.EventHandler(this.readFCToolStripMenuItem_Click);
            // 
            // readFCPlanningToolStripMenuItem
            // 
            this.readFCPlanningToolStripMenuItem.Name = "readFCPlanningToolStripMenuItem";
            this.readFCPlanningToolStripMenuItem.Size = new System.Drawing.Size(268, 34);
            this.readFCPlanningToolStripMenuItem.Text = "Read FC Planning";
            this.readFCPlanningToolStripMenuItem.Click += new System.EventHandler(this.readFCPlanningToolStripMenuItem_Click);
            // 
            // readOpenConfigToolStripMenuItem
            // 
            this.readOpenConfigToolStripMenuItem.Name = "readOpenConfigToolStripMenuItem";
            this.readOpenConfigToolStripMenuItem.Size = new System.Drawing.Size(268, 34);
            this.readOpenConfigToolStripMenuItem.Text = "Read OpenConfig";
            this.readOpenConfigToolStripMenuItem.Click += new System.EventHandler(this.readOpenConfigToolStripMenuItem_Click);
            // 
            // readDeliToolStripMenuItem
            // 
            this.readDeliToolStripMenuItem.Name = "readDeliToolStripMenuItem";
            this.readDeliToolStripMenuItem.Size = new System.Drawing.Size(268, 34);
            this.readDeliToolStripMenuItem.Text = "Read Deli";
            this.readDeliToolStripMenuItem.Click += new System.EventHandler(this.readDeliToolStripMenuItem_Click);
            // 
            // mastahToolStripMenuItem
            // 
            this.mastahToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fiteMoiToolStripMenuItem,
            this.noSupToolStripMenuItem,
            this.compactMastahToolStripMenuItem});
            this.mastahToolStripMenuItem.Name = "mastahToolStripMenuItem";
            this.mastahToolStripMenuItem.Size = new System.Drawing.Size(94, 34);
            this.mastahToolStripMenuItem.Text = "Mastah";
            // 
            // fiteMoiToolStripMenuItem
            // 
            this.fiteMoiToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.kgToolStripMenuItem,
            this.unitToolStripMenuItem});
            this.fiteMoiToolStripMenuItem.Name = "fiteMoiToolStripMenuItem";
            this.fiteMoiToolStripMenuItem.Size = new System.Drawing.Size(262, 34);
            this.fiteMoiToolStripMenuItem.Text = "Fite moi!";
            // 
            // kgToolStripMenuItem
            // 
            this.kgToolStripMenuItem.Name = "kgToolStripMenuItem";
            this.kgToolStripMenuItem.Size = new System.Drawing.Size(142, 34);
            this.kgToolStripMenuItem.Text = "Kg";
            this.kgToolStripMenuItem.Click += new System.EventHandler(this.kgToolStripMenuItem_Click);
            // 
            // unitToolStripMenuItem
            // 
            this.unitToolStripMenuItem.Name = "unitToolStripMenuItem";
            this.unitToolStripMenuItem.Size = new System.Drawing.Size(142, 34);
            this.unitToolStripMenuItem.Text = "Unit";
            this.unitToolStripMenuItem.Click += new System.EventHandler(this.unitToolStripMenuItem_Click);
            // 
            // noSupToolStripMenuItem
            // 
            this.noSupToolStripMenuItem.Name = "noSupToolStripMenuItem";
            this.noSupToolStripMenuItem.Size = new System.Drawing.Size(262, 34);
            this.noSupToolStripMenuItem.Text = "NoSup";
            this.noSupToolStripMenuItem.Click += new System.EventHandler(this.noSupToolStripMenuItem_Click);
            // 
            // compactMastahToolStripMenuItem
            // 
            this.compactMastahToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cap100ToolStripMenuItem,
            this.noCapToolStripMenuItem});
            this.compactMastahToolStripMenuItem.Name = "compactMastahToolStripMenuItem";
            this.compactMastahToolStripMenuItem.Size = new System.Drawing.Size(262, 34);
            this.compactMastahToolStripMenuItem.Text = "Compact Mastah";
            // 
            // cap100ToolStripMenuItem
            // 
            this.cap100ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.groupAllToolStripMenuItem,
            this.seperateFarmsToolStripMenuItem,
            this.seperateThuMuaToolStripMenuItem,
            this.seperateAllToolStripMenuItem,
            this.seperateAllOnlyFarmToolStripMenuItem});
            this.cap100ToolStripMenuItem.Name = "cap100ToolStripMenuItem";
            this.cap100ToolStripMenuItem.Size = new System.Drawing.Size(196, 34);
            this.cap100ToolStripMenuItem.Text = "Cap 100%";
            // 
            // groupAllToolStripMenuItem
            // 
            this.groupAllToolStripMenuItem.Name = "groupAllToolStripMenuItem";
            this.groupAllToolStripMenuItem.Size = new System.Drawing.Size(315, 34);
            this.groupAllToolStripMenuItem.Text = "Group All";
            this.groupAllToolStripMenuItem.Click += new System.EventHandler(this.groupAllToolStripMenuItem_Click);
            // 
            // seperateFarmsToolStripMenuItem
            // 
            this.seperateFarmsToolStripMenuItem.Name = "seperateFarmsToolStripMenuItem";
            this.seperateFarmsToolStripMenuItem.Size = new System.Drawing.Size(315, 34);
            this.seperateFarmsToolStripMenuItem.Text = "Seperate Farms";
            this.seperateFarmsToolStripMenuItem.Click += new System.EventHandler(this.seperateFarmsToolStripMenuItem_Click);
            // 
            // seperateThuMuaToolStripMenuItem
            // 
            this.seperateThuMuaToolStripMenuItem.Name = "seperateThuMuaToolStripMenuItem";
            this.seperateThuMuaToolStripMenuItem.Size = new System.Drawing.Size(315, 34);
            this.seperateThuMuaToolStripMenuItem.Text = "Seperate ThuMua";
            this.seperateThuMuaToolStripMenuItem.Click += new System.EventHandler(this.seperateThuMuaToolStripMenuItem_Click);
            // 
            // seperateAllToolStripMenuItem
            // 
            this.seperateAllToolStripMenuItem.Name = "seperateAllToolStripMenuItem";
            this.seperateAllToolStripMenuItem.Size = new System.Drawing.Size(315, 34);
            this.seperateAllToolStripMenuItem.Text = "Seperate All";
            this.seperateAllToolStripMenuItem.Click += new System.EventHandler(this.seperateAllToolStripMenuItem_Click);
            // 
            // seperateAllOnlyFarmToolStripMenuItem
            // 
            this.seperateAllOnlyFarmToolStripMenuItem.Name = "seperateAllOnlyFarmToolStripMenuItem";
            this.seperateAllOnlyFarmToolStripMenuItem.Size = new System.Drawing.Size(315, 34);
            this.seperateAllOnlyFarmToolStripMenuItem.Text = "Seperate All Only Farm";
            this.seperateAllOnlyFarmToolStripMenuItem.Click += new System.EventHandler(this.seperateAllOnlyFarmToolStripMenuItem_Click);
            // 
            // noCapToolStripMenuItem
            // 
            this.noCapToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.groupAllNoCapToolStripMenuItem,
            this.seperateFarmNoCapToolStripMenuItem,
            this.seperateThuMuaNoCapToolStripMenuItem,
            this.seperateAllNoCapToolStripMenuItem,
            this.seperateAllOnlyFarmToolStripMenuItem1});
            this.noCapToolStripMenuItem.Name = "noCapToolStripMenuItem";
            this.noCapToolStripMenuItem.Size = new System.Drawing.Size(196, 34);
            this.noCapToolStripMenuItem.Text = "NoCap";
            // 
            // groupAllNoCapToolStripMenuItem
            // 
            this.groupAllNoCapToolStripMenuItem.Name = "groupAllNoCapToolStripMenuItem";
            this.groupAllNoCapToolStripMenuItem.Size = new System.Drawing.Size(338, 34);
            this.groupAllNoCapToolStripMenuItem.Text = "Group All NoCap";
            this.groupAllNoCapToolStripMenuItem.Click += new System.EventHandler(this.groupAllNoCapToolStripMenuItem_Click);
            // 
            // seperateFarmNoCapToolStripMenuItem
            // 
            this.seperateFarmNoCapToolStripMenuItem.Name = "seperateFarmNoCapToolStripMenuItem";
            this.seperateFarmNoCapToolStripMenuItem.Size = new System.Drawing.Size(338, 34);
            this.seperateFarmNoCapToolStripMenuItem.Text = "Seperate Farm NoCap";
            this.seperateFarmNoCapToolStripMenuItem.Click += new System.EventHandler(this.seperateFarmNoCapToolStripMenuItem_Click);
            // 
            // seperateThuMuaNoCapToolStripMenuItem
            // 
            this.seperateThuMuaNoCapToolStripMenuItem.Name = "seperateThuMuaNoCapToolStripMenuItem";
            this.seperateThuMuaNoCapToolStripMenuItem.Size = new System.Drawing.Size(338, 34);
            this.seperateThuMuaNoCapToolStripMenuItem.Text = "Seperate ThuMua NoCap";
            this.seperateThuMuaNoCapToolStripMenuItem.Click += new System.EventHandler(this.seperateThuMuaNoCapToolStripMenuItem_Click);
            // 
            // seperateAllNoCapToolStripMenuItem
            // 
            this.seperateAllNoCapToolStripMenuItem.Name = "seperateAllNoCapToolStripMenuItem";
            this.seperateAllNoCapToolStripMenuItem.Size = new System.Drawing.Size(338, 34);
            this.seperateAllNoCapToolStripMenuItem.Text = "Seperate All NoCap";
            this.seperateAllNoCapToolStripMenuItem.Click += new System.EventHandler(this.seperateAllNoCapToolStripMenuItem_Click);
            // 
            // seperateAllOnlyFarmToolStripMenuItem1
            // 
            this.seperateAllOnlyFarmToolStripMenuItem1.Name = "seperateAllOnlyFarmToolStripMenuItem1";
            this.seperateAllOnlyFarmToolStripMenuItem1.Size = new System.Drawing.Size(338, 34);
            this.seperateAllOnlyFarmToolStripMenuItem1.Text = "Seperate All Only Farm";
            this.seperateAllOnlyFarmToolStripMenuItem1.Click += new System.EventHandler(this.seperateAllOnlyFarmToolStripMenuItem1_Click);
            // 
            // printToolStripMenuItem
            // 
            this.printToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.purchaseOrderToolStripMenuItem});
            this.printToolStripMenuItem.Name = "printToolStripMenuItem";
            this.printToolStripMenuItem.Size = new System.Drawing.Size(68, 34);
            this.printToolStripMenuItem.Text = "Print";
            // 
            // purchaseOrderToolStripMenuItem
            // 
            this.purchaseOrderToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.horizontallyToolStripMenuItem,
            this.verticallyToolStripMenuItem,
            this.compactToolStripMenuItem});
            this.purchaseOrderToolStripMenuItem.Name = "purchaseOrderToolStripMenuItem";
            this.purchaseOrderToolStripMenuItem.Size = new System.Drawing.Size(247, 34);
            this.purchaseOrderToolStripMenuItem.Text = "Purchase Order";
            // 
            // horizontallyToolStripMenuItem
            // 
            this.horizontallyToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.unitToolStripMenuItem1,
            this.kgToolStripMenuItem1});
            this.horizontallyToolStripMenuItem.Name = "horizontallyToolStripMenuItem";
            this.horizontallyToolStripMenuItem.Size = new System.Drawing.Size(215, 34);
            this.horizontallyToolStripMenuItem.Text = "Horizontally";
            // 
            // unitToolStripMenuItem1
            // 
            this.unitToolStripMenuItem1.Name = "unitToolStripMenuItem1";
            this.unitToolStripMenuItem1.Size = new System.Drawing.Size(142, 34);
            this.unitToolStripMenuItem1.Text = "Unit";
            this.unitToolStripMenuItem1.Click += new System.EventHandler(this.unitToolStripMenuItem1_Click);
            // 
            // kgToolStripMenuItem1
            // 
            this.kgToolStripMenuItem1.Name = "kgToolStripMenuItem1";
            this.kgToolStripMenuItem1.Size = new System.Drawing.Size(142, 34);
            this.kgToolStripMenuItem1.Text = "Kg";
            this.kgToolStripMenuItem1.Click += new System.EventHandler(this.kgToolStripMenuItem1_Click);
            // 
            // verticallyToolStripMenuItem
            // 
            this.verticallyToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.unitToolStripMenuItem2,
            this.kgToolStripMenuItem2});
            this.verticallyToolStripMenuItem.Name = "verticallyToolStripMenuItem";
            this.verticallyToolStripMenuItem.Size = new System.Drawing.Size(215, 34);
            this.verticallyToolStripMenuItem.Text = "Vertically";
            // 
            // unitToolStripMenuItem2
            // 
            this.unitToolStripMenuItem2.Name = "unitToolStripMenuItem2";
            this.unitToolStripMenuItem2.Size = new System.Drawing.Size(142, 34);
            this.unitToolStripMenuItem2.Text = "Unit";
            this.unitToolStripMenuItem2.Click += new System.EventHandler(this.unitToolStripMenuItem2_Click);
            // 
            // kgToolStripMenuItem2
            // 
            this.kgToolStripMenuItem2.Name = "kgToolStripMenuItem2";
            this.kgToolStripMenuItem2.Size = new System.Drawing.Size(142, 34);
            this.kgToolStripMenuItem2.Text = "Kg";
            this.kgToolStripMenuItem2.Click += new System.EventHandler(this.kgToolStripMenuItem2_Click);
            // 
            // compactToolStripMenuItem
            // 
            this.compactToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.compactPOunitToolStripMenuItem,
            this.compactPOkgToolStripMenuItem,
            this.btnPrintPOReport});
            this.compactToolStripMenuItem.Name = "compactToolStripMenuItem";
            this.compactToolStripMenuItem.Size = new System.Drawing.Size(215, 34);
            this.compactToolStripMenuItem.Text = "Compact";
            // 
            // compactPOunitToolStripMenuItem
            // 
            this.compactPOunitToolStripMenuItem.Name = "compactPOunitToolStripMenuItem";
            this.compactPOunitToolStripMenuItem.Size = new System.Drawing.Size(165, 34);
            this.compactPOunitToolStripMenuItem.Text = "Unit";
            this.compactPOunitToolStripMenuItem.Click += new System.EventHandler(this.compactPOunitToolStripMenuItem_Click);
            // 
            // compactPOkgToolStripMenuItem
            // 
            this.compactPOkgToolStripMenuItem.Name = "compactPOkgToolStripMenuItem";
            this.compactPOkgToolStripMenuItem.Size = new System.Drawing.Size(165, 34);
            this.compactPOkgToolStripMenuItem.Text = "Kg";
            this.compactPOkgToolStripMenuItem.Click += new System.EventHandler(this.compactPOkgToolStripMenuItem_Click);
            // 
            // btnPrintPOReport
            // 
            this.btnPrintPOReport.Name = "btnPrintPOReport";
            this.btnPrintPOReport.Size = new System.Drawing.Size(165, 34);
            this.btnPrintPOReport.Text = "Report";
            this.btnPrintPOReport.Click += new System.EventHandler(this.btnPrintPOReport_Click);
            // 
            // testStuffToolStripMenuItem
            // 
            this.testStuffToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.testExportLargeExcelToolStripMenuItem});
            this.testStuffToolStripMenuItem.Name = "testStuffToolStripMenuItem";
            this.testStuffToolStripMenuItem.Size = new System.Drawing.Size(109, 34);
            this.testStuffToolStripMenuItem.Text = "Test stuff";
            this.testStuffToolStripMenuItem.Click += new System.EventHandler(this.testExportLargeExcelToolStripMenuItem_Click);
            // 
            // testExportLargeExcelToolStripMenuItem
            // 
            this.testExportLargeExcelToolStripMenuItem.Name = "testExportLargeExcelToolStripMenuItem";
            this.testExportLargeExcelToolStripMenuItem.Size = new System.Drawing.Size(315, 34);
            this.testExportLargeExcelToolStripMenuItem.Text = "Test Export Large Excel";
            // 
            // reportToolStripMenuItem
            // 
            this.reportToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.m1ToolStripMenuItem});
            this.reportToolStripMenuItem.Name = "reportToolStripMenuItem";
            this.reportToolStripMenuItem.Size = new System.Drawing.Size(86, 34);
            this.reportToolStripMenuItem.Text = "Report";
            // 
            // m1ToolStripMenuItem
            // 
            this.m1ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.printToolStripMenuItem1});
            this.m1ToolStripMenuItem.Name = "m1ToolStripMenuItem";
            this.m1ToolStripMenuItem.Size = new System.Drawing.Size(148, 34);
            this.m1ToolStripMenuItem.Text = "M+1";
            // 
            // printToolStripMenuItem1
            // 
            this.printToolStripMenuItem1.Name = "printToolStripMenuItem1";
            this.printToolStripMenuItem1.Size = new System.Drawing.Size(147, 34);
            this.printToolStripMenuItem1.Text = "Print";
            this.printToolStripMenuItem1.Click += new System.EventHandler(this.printToolStripMenuItem1_Click);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mongoDbToolStripMenuItem});
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            this.deleteToolStripMenuItem.Size = new System.Drawing.Size(85, 34);
            this.deleteToolStripMenuItem.Text = "Delete";
            // 
            // mongoDbToolStripMenuItem
            // 
            this.mongoDbToolStripMenuItem.Name = "mongoDbToolStripMenuItem";
            this.mongoDbToolStripMenuItem.Size = new System.Drawing.Size(198, 34);
            this.mongoDbToolStripMenuItem.Text = "MongoDb";
            this.mongoDbToolStripMenuItem.Click += new System.EventHandler(this.mongoDbToolStripMenuItem_Click);
            // 
            // DateFromPicker
            // 
            this.DateFromPicker.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DateFromPicker.Location = new System.Drawing.Point(38, 65);
            this.DateFromPicker.Name = "DateFromPicker";
            this.DateFromPicker.Size = new System.Drawing.Size(200, 29);
            this.DateFromPicker.TabIndex = 3;
            this.DateFromPicker.ValueChanged += new System.EventHandler(this.DateFromPicker_ValueChanged);
            // 
            // DateToPicker
            // 
            this.DateToPicker.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DateToPicker.Location = new System.Drawing.Point(339, 65);
            this.DateToPicker.Name = "DateToPicker";
            this.DateToPicker.Size = new System.Drawing.Size(200, 29);
            this.DateToPicker.TabIndex = 4;
            this.DateToPicker.ValueChanged += new System.EventHandler(this.DateToPicker_ValueChanged);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(38, 143);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(501, 41);
            this.progressBar1.TabIndex = 5;
            // 
            // progressBarLabel
            // 
            this.progressBarLabel.AutoSize = true;
            this.progressBarLabel.Location = new System.Drawing.Point(33, 143);
            this.progressBarLabel.Name = "progressBarLabel";
            this.progressBarLabel.Size = new System.Drawing.Size(64, 25);
            this.progressBarLabel.TabIndex = 6;
            this.progressBarLabel.Text = "label1";
            this.progressBarLabel.Visible = false;
            // 
            // upperCapBox
            // 
            this.upperCapBox.Location = new System.Drawing.Point(575, 65);
            this.upperCapBox.Name = "upperCapBox";
            this.upperCapBox.Size = new System.Drawing.Size(64, 29);
            this.upperCapBox.TabIndex = 7;
            this.upperCapBox.TextChanged += new System.EventHandler(this.upperCapBox_TextChanged);
            // 
            // checkBoxFruitOnly
            // 
            this.checkBoxFruitOnly.AutoSize = true;
            this.checkBoxFruitOnly.CheckAlign = System.Drawing.ContentAlignment.TopLeft;
            this.checkBoxFruitOnly.Location = new System.Drawing.Point(672, 100);
            this.checkBoxFruitOnly.Name = "checkBoxFruitOnly";
            this.checkBoxFruitOnly.Size = new System.Drawing.Size(128, 29);
            this.checkBoxFruitOnly.TabIndex = 8;
            this.checkBoxFruitOnly.Text = "Fruit Only!";
            this.checkBoxFruitOnly.UseVisualStyleBackColor = true;
            this.checkBoxFruitOnly.CheckedChanged += new System.EventHandler(this.checkBoxFruitOnly_CheckedChanged);
            // 
            // checkBoxNoFruit
            // 
            this.checkBoxNoFruit.AutoSize = true;
            this.checkBoxNoFruit.CheckAlign = System.Drawing.ContentAlignment.TopLeft;
            this.checkBoxNoFruit.Location = new System.Drawing.Point(672, 65);
            this.checkBoxNoFruit.Name = "checkBoxNoFruit";
            this.checkBoxNoFruit.Size = new System.Drawing.Size(112, 29);
            this.checkBoxNoFruit.TabIndex = 9;
            this.checkBoxNoFruit.Text = "No Fruit!";
            this.checkBoxNoFruit.UseVisualStyleBackColor = true;
            this.checkBoxNoFruit.CheckedChanged += new System.EventHandler(this.checkBoxNoFruit_CheckedChanged);
            // 
            // YesNoSubRegionchkBox
            // 
            this.YesNoSubRegionchkBox.AutoSize = true;
            this.YesNoSubRegionchkBox.Location = new System.Drawing.Point(672, 135);
            this.YesNoSubRegionchkBox.Name = "YesNoSubRegionchkBox";
            this.YesNoSubRegionchkBox.Size = new System.Drawing.Size(191, 29);
            this.YesNoSubRegionchkBox.TabIndex = 10;
            this.YesNoSubRegionchkBox.Text = "With Sub Region!";
            this.YesNoSubRegionchkBox.UseVisualStyleBackColor = true;
            this.YesNoSubRegionchkBox.CheckedChanged += new System.EventHandler(this.YesNoSubRegion_CheckedChanged);
            // 
            // richTextBoxOutput
            // 
            this.richTextBoxOutput.Font = new System.Drawing.Font("Calibri", 11.14286F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.richTextBoxOutput.Location = new System.Drawing.Point(12, 307);
            this.richTextBoxOutput.Name = "richTextBoxOutput";
            this.richTextBoxOutput.Size = new System.Drawing.Size(1160, 449);
            this.richTextBoxOutput.TabIndex = 11;
            this.richTextBoxOutput.Text = "";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::AllocatingStuff.Properties.Resources.b;
            this.ClientSize = new System.Drawing.Size(1184, 768);
            this.Controls.Add(this.richTextBoxOutput);
            this.Controls.Add(this.YesNoSubRegionchkBox);
            this.Controls.Add(this.checkBoxNoFruit);
            this.Controls.Add(this.checkBoxFruitOnly);
            this.Controls.Add(this.upperCapBox);
            this.Controls.Add(this.progressBarLabel);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.DateToPicker);
            this.Controls.Add(this.DateFromPicker);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.Text = "ChiaHang";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ToolStripMenuItem navigateToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem readPOToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mastahToolStripMenuItem;
        private System.Windows.Forms.DateTimePicker DateFromPicker;
        private System.Windows.Forms.DateTimePicker DateToPicker;
        private System.Windows.Forms.ToolStripMenuItem readFCToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fiteMoiToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem noSupToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem printToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem purchaseOrderToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem testStuffToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem testExportLargeExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem compactMastahToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cap100ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem noCapToolStripMenuItem;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label progressBarLabel;
        private System.Windows.Forms.ToolStripMenuItem seperateFarmsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem groupAllNoCapToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem seperateFarmNoCapToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem seperateThuMuaNoCapToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem seperateAllNoCapToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem groupAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem seperateThuMuaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem seperateAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem readFCPlanningToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem horizontallyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem verticallyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem reportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem m1ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem printToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem kgToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem unitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem readOpenConfigToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem unitToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem kgToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem unitToolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem kgToolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem seperateAllOnlyFarmToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deleteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mongoDbToolStripMenuItem;
        private System.Windows.Forms.TextBox upperCapBox;
        private System.Windows.Forms.ToolStripMenuItem seperateAllOnlyFarmToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem compactToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem compactPOunitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem compactPOkgToolStripMenuItem;
        private System.Windows.Forms.CheckBox checkBoxFruitOnly;
        private System.Windows.Forms.CheckBox checkBoxNoFruit;
        private System.Windows.Forms.ToolStripMenuItem btnPrintPOReport;
        private System.Windows.Forms.ToolStripMenuItem readDeliToolStripMenuItem;
        private System.Windows.Forms.CheckBox YesNoSubRegionchkBox;
        private System.Windows.Forms.RichTextBox richTextBoxOutput;
        private System.ComponentModel.BackgroundWorker backgroundCoord;
    }
}

