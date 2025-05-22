namespace UygunOlmayan
{
    partial class RAPOR
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RAPOR));
            advancedDataGridView1 = new Zuby.ADGV.AdvancedDataGridView();
            menuStrip1 = new MenuStrip();
            fORMToolStripMenuItem = new ToolStripMenuItem();
            lİSTEYİEXCELEAKTARToolStripMenuItem = new ToolStripMenuItem();
            gERİDÖNToolStripMenuItem = new ToolStripMenuItem();
            lİSTEToolStripMenuItem = new ToolStripMenuItem();
            groupBox3 = new GroupBox();
            comboBox3 = new ComboBox();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).BeginInit();
            menuStrip1.SuspendLayout();
            groupBox3.SuspendLayout();
            SuspendLayout();
            // 
            // advancedDataGridView1
            // 
            advancedDataGridView1.AllowUserToAddRows = false;
            advancedDataGridView1.AllowUserToDeleteRows = false;
            advancedDataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            advancedDataGridView1.BackgroundColor = Color.White;
            advancedDataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            advancedDataGridView1.FilterAndSortEnabled = true;
            advancedDataGridView1.FilterStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.Location = new Point(30, 140);
            advancedDataGridView1.Name = "advancedDataGridView1";
            advancedDataGridView1.ReadOnly = true;
            advancedDataGridView1.RightToLeft = RightToLeft.No;
            advancedDataGridView1.RowHeadersWidth = 51;
            advancedDataGridView1.RowTemplate.Height = 29;
            advancedDataGridView1.Size = new Size(1843, 839);
            advancedDataGridView1.SortStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.TabIndex = 4;
            // 
            // menuStrip1
            // 
            menuStrip1.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            menuStrip1.ImageScalingSize = new Size(20, 20);
            menuStrip1.Items.AddRange(new ToolStripItem[] { fORMToolStripMenuItem, lİSTEToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(1902, 28);
            menuStrip1.TabIndex = 54;
            menuStrip1.Text = "menuStrip1";
            // 
            // fORMToolStripMenuItem
            // 
            fORMToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { lİSTEYİEXCELEAKTARToolStripMenuItem, gERİDÖNToolStripMenuItem });
            fORMToolStripMenuItem.Image = (Image)resources.GetObject("fORMToolStripMenuItem.Image");
            fORMToolStripMenuItem.Name = "fORMToolStripMenuItem";
            fORMToolStripMenuItem.Size = new Size(92, 24);
            fORMToolStripMenuItem.Text = "DOSYA";
            // 
            // lİSTEYİEXCELEAKTARToolStripMenuItem
            // 
            lİSTEYİEXCELEAKTARToolStripMenuItem.Image = (Image)resources.GetObject("lİSTEYİEXCELEAKTARToolStripMenuItem.Image");
            lİSTEYİEXCELEAKTARToolStripMenuItem.Name = "lİSTEYİEXCELEAKTARToolStripMenuItem";
            lİSTEYİEXCELEAKTARToolStripMenuItem.Size = new Size(253, 26);
            lİSTEYİEXCELEAKTARToolStripMenuItem.Text = "LİSTEYİ EXCELE AKTAR";
            lİSTEYİEXCELEAKTARToolStripMenuItem.Click += lİSTEYİEXCELEAKTARToolStripMenuItem_Click;
            // 
            // gERİDÖNToolStripMenuItem
            // 
            gERİDÖNToolStripMenuItem.Image = (Image)resources.GetObject("gERİDÖNToolStripMenuItem.Image");
            gERİDÖNToolStripMenuItem.Name = "gERİDÖNToolStripMenuItem";
            gERİDÖNToolStripMenuItem.Size = new Size(253, 26);
            gERİDÖNToolStripMenuItem.Text = "GERİ DÖN";
            // 
            // lİSTEToolStripMenuItem
            // 
            lİSTEToolStripMenuItem.Name = "lİSTEToolStripMenuItem";
            lİSTEToolStripMenuItem.Size = new Size(14, 24);
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(comboBox3);
            groupBox3.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            groupBox3.Location = new Point(30, 40);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(280, 62);
            groupBox3.TabIndex = 55;
            groupBox3.TabStop = false;
            groupBox3.Text = "Hatanı Oluştuğu Bölümü";
            // 
            // comboBox3
            // 
            comboBox3.FormattingEnabled = true;
            comboBox3.Items.AddRange(new object[] { "Montaj", "Tasarım", "İmalat", "Otomasyon", "Satınalma", "Planlama", "Kalite Kontrol", "Fabrika Müdürü", "Satış Sonrası", "Muhasebe" });
            comboBox3.Location = new Point(6, 26);
            comboBox3.Name = "comboBox3";
            comboBox3.Size = new Size(262, 28);
            comboBox3.TabIndex = 10;
            comboBox3.Text = "Hata Bölümü seçiniz...";
            comboBox3.SelectedIndexChanged += comboBox3_SelectedIndexChanged;
            // 
            // RAPOR
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Gray;
            ClientSize = new Size(1902, 1033);
            Controls.Add(groupBox3);
            Controls.Add(menuStrip1);
            Controls.Add(advancedDataGridView1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "RAPOR";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "RAPOR";
            WindowState = FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).EndInit();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            groupBox3.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Zuby.ADGV.AdvancedDataGridView advancedDataGridView1;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fORMToolStripMenuItem;
        private ToolStripMenuItem lİSTEYİEXCELEAKTARToolStripMenuItem;
        private ToolStripMenuItem gERİDÖNToolStripMenuItem;
        private ToolStripMenuItem lİSTEToolStripMenuItem;
        private GroupBox groupBox3;
        private ComboBox comboBox3;
    }
}