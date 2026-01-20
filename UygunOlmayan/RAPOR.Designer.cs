namespace UygunOlmayan
{
    partial class RAPOR
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle3 = new DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RAPOR));
            advancedDataGridView1 = new Zuby.ADGV.AdvancedDataGridView();
            panelHeader = new Panel();
            lblTitle = new Label();
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
            advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            advancedDataGridView1.BackgroundColor = Color.White;
            advancedDataGridView1.BorderStyle = BorderStyle.None;
            advancedDataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            advancedDataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = Color.FromArgb(41, 57, 85);
            dataGridViewCellStyle1.Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold);
            dataGridViewCellStyle1.ForeColor = Color.White;
            dataGridViewCellStyle1.SelectionBackColor = Color.FromArgb(41, 57, 85);
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
            advancedDataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            advancedDataGridView1.ColumnHeadersHeight = 45;
            advancedDataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = Color.White;
            dataGridViewCellStyle2.Font = new Font("Segoe UI", 9F);
            dataGridViewCellStyle2.ForeColor = Color.FromArgb(64, 64, 64);
            dataGridViewCellStyle2.SelectionBackColor = Color.FromArgb(226, 230, 236);
            dataGridViewCellStyle2.SelectionForeColor = Color.Black;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.False;
            advancedDataGridView1.DefaultCellStyle = dataGridViewCellStyle2;
            advancedDataGridView1.EnableHeadersVisualStyles = false;
            advancedDataGridView1.FilterAndSortEnabled = true;
            advancedDataGridView1.FilterStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.GridColor = Color.FromArgb(224, 224, 224);
            advancedDataGridView1.Location = new Point(12, 190);
            advancedDataGridView1.Name = "advancedDataGridView1";
            advancedDataGridView1.ReadOnly = true;
            advancedDataGridView1.RowHeadersVisible = false;
            advancedDataGridView1.RowHeadersWidth = 51;
            dataGridViewCellStyle3.BackColor = Color.FromArgb(250, 251, 252);
            advancedDataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            advancedDataGridView1.RowTemplate.Height = 40;
            advancedDataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            advancedDataGridView1.Size = new Size(1878, 831);
            advancedDataGridView1.TabIndex = 4;
            advancedDataGridView1.SortStringChangedInvokeBeforeDatasourceUpdate = true;
            // 
            // panelHeader
            // 
            panelHeader.BackColor = Color.FromArgb(41, 57, 85);
            panelHeader.Controls.Add(lblTitle);
            panelHeader.Dock = DockStyle.Top;
            panelHeader.Location = new Point(0, 28);
            panelHeader.Name = "panelHeader";
            panelHeader.Size = new Size(1902, 80);
            panelHeader.TabIndex = 56;
            // 
            // lblTitle
            // 
            lblTitle.AutoSize = true;
            lblTitle.Font = new Font("Segoe UI", 20F, FontStyle.Bold);
            lblTitle.ForeColor = Color.White;
            lblTitle.Location = new Point(20, 18);
            lblTitle.Name = "lblTitle";
            lblTitle.Size = new Size(331, 46);
            lblTitle.TabIndex = 0;
            lblTitle.Text = "ANALİZ RAPORLARI";
            // 
            // menuStrip1
            // 
            menuStrip1.BackColor = Color.FromArgb(41, 57, 85);
            menuStrip1.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            menuStrip1.ForeColor = Color.White;
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
            fORMToolStripMenuItem.ForeColor = Color.White;
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
            groupBox3.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            groupBox3.ForeColor = Color.FromArgb(64, 64, 64);
            groupBox3.Location = new Point(12, 115);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(350, 65);
            groupBox3.TabIndex = 55;
            groupBox3.TabStop = false;
            groupBox3.Text = "BÖLÜM FİLTRESİ";
            // 
            // comboBox3
            // 
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox3.FlatStyle = FlatStyle.Flat;
            comboBox3.FormattingEnabled = true;
            comboBox3.Items.AddRange(new object[] { "TÜMÜ", "Montaj", "Tasarım", "İmalat", "Otomasyon", "Satınalma", "Planlama", "Kalite Kontrol", "Fabrika Müdürü", "Satış Sonrası", "Muhasebe" });
            comboBox3.Location = new Point(10, 25);
            comboBox3.Name = "comboBox3";
            comboBox3.Size = new Size(330, 28);
            comboBox3.TabIndex = 10;
            comboBox3.SelectedIndexChanged += comboBox3_SelectedIndexChanged;
            // 
            // RAPOR
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(242, 245, 250);
            ClientSize = new Size(1902, 1033);
            Controls.Add(groupBox3);
            Controls.Add(panelHeader);
            Controls.Add(menuStrip1);
            Controls.Add(advancedDataGridView1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "RAPOR";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Analiz Raporları";
            WindowState = FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).EndInit();
            panelHeader.ResumeLayout(false);
            panelHeader.PerformLayout();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            groupBox3.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Zuby.ADGV.AdvancedDataGridView advancedDataGridView1;
        private Panel panelHeader;
        private Label lblTitle;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fORMToolStripMenuItem;
        private ToolStripMenuItem lİSTEYİEXCELEAKTARToolStripMenuItem;
        private ToolStripMenuItem gERİDÖNToolStripMenuItem;
        private ToolStripMenuItem lİSTEToolStripMenuItem;
        private GroupBox groupBox3;
        private ComboBox comboBox3;
    }
}