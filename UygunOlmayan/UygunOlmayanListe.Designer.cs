namespace UygunOlmayan
{
    partial class UygunOlmayanListe
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UygunOlmayanListe));
            advancedDataGridView1 = new Zuby.ADGV.AdvancedDataGridView();
            panelHeader = new Panel();
            lblTitle = new Label();
            menuStrip1 = new MenuStrip();
            fORMToolStripMenuItem = new ToolStripMenuItem();
            eXCELÇEKToolStripMenuItem = new ToolStripMenuItem();
            ePOSTAGÖNDERToolStripMenuItem = new ToolStripMenuItem();
            lİSTEYİEXCELEAKTARToolStripMenuItem = new ToolStripMenuItem();
            gERİDÖNToolStripMenuItem = new ToolStripMenuItem();
            lİSTEToolStripMenuItem = new ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).BeginInit();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // advancedDataGridView1
            // 
            advancedDataGridView1.AllowUserToAddRows = false;
            advancedDataGridView1.AllowUserToDeleteRows = false;
            advancedDataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            advancedDataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
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
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.True;
            advancedDataGridView1.DefaultCellStyle = dataGridViewCellStyle2;
            advancedDataGridView1.EnableHeadersVisualStyles = false;
            advancedDataGridView1.FilterAndSortEnabled = true;
            advancedDataGridView1.FilterStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.GridColor = Color.FromArgb(224, 224, 224);
            advancedDataGridView1.Location = new Point(12, 120);
            advancedDataGridView1.Name = "advancedDataGridView1";
            advancedDataGridView1.ReadOnly = true;
            advancedDataGridView1.RowHeadersVisible = false;
            advancedDataGridView1.RowHeadersWidth = 51;
            dataGridViewCellStyle3.BackColor = Color.FromArgb(250, 251, 252);
            advancedDataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            advancedDataGridView1.RowTemplate.Height = 40;
            advancedDataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            advancedDataGridView1.Size = new Size(1878, 901);
            advancedDataGridView1.TabIndex = 3;
            advancedDataGridView1.FilterStringChanged += advancedDataGridView1_FilterStringChanged;
            advancedDataGridView1.CellContentClick += advancedDataGridView1_CellContentClick;
            advancedDataGridView1.CellDoubleClick += advancedDataGridView1_CellDoubleClick;
            advancedDataGridView1.DataBindingComplete += advancedDataGridView1_DataBindingComplete;
            // 
            // panelHeader
            // 
            panelHeader.BackColor = Color.FromArgb(41, 57, 85);
            panelHeader.Controls.Add(lblTitle);
            panelHeader.Dock = DockStyle.Top;
            panelHeader.Location = new Point(0, 28);
            panelHeader.Name = "panelHeader";
            panelHeader.Size = new Size(1902, 80);
            panelHeader.TabIndex = 54;
            // 
            // lblTitle
            // 
            lblTitle.AutoSize = true;
            lblTitle.Font = new Font("Segoe UI", 20F, FontStyle.Bold);
            lblTitle.ForeColor = Color.White;
            lblTitle.Location = new Point(20, 18);
            lblTitle.Name = "lblTitle";
            lblTitle.Size = new Size(580, 46);
            lblTitle.TabIndex = 0;
            lblTitle.Text = "UYGUN OLMAYAN ÜRÜN LİSTESİ";
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
            menuStrip1.TabIndex = 53;
            menuStrip1.Text = "menuStrip1";
            // 
            // fORMToolStripMenuItem
            // 
            fORMToolStripMenuItem.ForeColor = Color.White;
            fORMToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { eXCELÇEKToolStripMenuItem, ePOSTAGÖNDERToolStripMenuItem, lİSTEYİEXCELEAKTARToolStripMenuItem, gERİDÖNToolStripMenuItem });
            fORMToolStripMenuItem.Image = (Image)resources.GetObject("fORMToolStripMenuItem.Image");
            fORMToolStripMenuItem.Name = "fORMToolStripMenuItem";
            fORMToolStripMenuItem.Size = new Size(92, 24);
            fORMToolStripMenuItem.Text = "DOSYA";
            // 
            // eXCELÇEKToolStripMenuItem
            // 
            eXCELÇEKToolStripMenuItem.Image = (Image)resources.GetObject("eXCELÇEKToolStripMenuItem.Image");
            eXCELÇEKToolStripMenuItem.Name = "eXCELÇEKToolStripMenuItem";
            eXCELÇEKToolStripMenuItem.Size = new Size(253, 26);
            eXCELÇEKToolStripMenuItem.Text = "EXCELE AKTAR";
            eXCELÇEKToolStripMenuItem.Click += eXCELÇEKToolStripMenuItem_Click;
            // 
            // ePOSTAGÖNDERToolStripMenuItem
            // 
            ePOSTAGÖNDERToolStripMenuItem.Image = (Image)resources.GetObject("ePOSTAGÖNDERToolStripMenuItem.Image");
            ePOSTAGÖNDERToolStripMenuItem.Name = "ePOSTAGÖNDERToolStripMenuItem";
            ePOSTAGÖNDERToolStripMenuItem.Size = new Size(253, 26);
            ePOSTAGÖNDERToolStripMenuItem.Text = "E-POSTA GÖNDER";
            ePOSTAGÖNDERToolStripMenuItem.Click += ePOSTAGÖNDERToolStripMenuItem_Click;
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
            gERİDÖNToolStripMenuItem.Click += gERİDÖNToolStripMenuItem_Click;
            // 
            // lİSTEToolStripMenuItem
            // 
            lİSTEToolStripMenuItem.Name = "lİSTEToolStripMenuItem";
            lİSTEToolStripMenuItem.Size = new Size(14, 24);
            // 
            // UygunOlmayanListe
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(242, 245, 250);
            ClientSize = new Size(1902, 1033);
            Controls.Add(panelHeader);
            Controls.Add(menuStrip1);
            Controls.Add(advancedDataGridView1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "UygunOlmayanListe";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Uygun Olmayan Ürün Listesi";
            WindowState = FormWindowState.Maximized;
            Load += UygunOlmayanListe_Load;
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).EndInit();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private Zuby.ADGV.AdvancedDataGridView advancedDataGridView1;
        private Panel panelHeader;
        private Label lblTitle;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fORMToolStripMenuItem;
        private ToolStripMenuItem eXCELÇEKToolStripMenuItem;
        private ToolStripMenuItem ePOSTAGÖNDERToolStripMenuItem;
        private ToolStripMenuItem lİSTEYİEXCELEAKTARToolStripMenuItem;
        private ToolStripMenuItem lİSTEToolStripMenuItem;
        private ToolStripMenuItem gERİDÖNToolStripMenuItem;
    }
}