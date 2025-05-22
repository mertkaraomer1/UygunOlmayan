namespace UygunOlmayan
{
    partial class UygunOlmayanListe
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
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UygunOlmayanListe));
            advancedDataGridView1 = new Zuby.ADGV.AdvancedDataGridView();
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
            advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader;
            advancedDataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            advancedDataGridView1.BackgroundColor = Color.White;
            advancedDataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = SystemColors.Window;
            dataGridViewCellStyle1.Font = new Font("Segoe UI", 9F);
            dataGridViewCellStyle1.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
            advancedDataGridView1.DefaultCellStyle = dataGridViewCellStyle1;
            advancedDataGridView1.FilterAndSortEnabled = true;
            advancedDataGridView1.FilterStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.Location = new Point(24, 71);
            advancedDataGridView1.Name = "advancedDataGridView1";
            advancedDataGridView1.ReadOnly = true;
            advancedDataGridView1.RightToLeft = RightToLeft.No;
            advancedDataGridView1.RowHeadersWidth = 51;
            advancedDataGridView1.RowTemplate.Height = 29;
            advancedDataGridView1.Size = new Size(1843, 925);
            advancedDataGridView1.SortStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.TabIndex = 3;
            advancedDataGridView1.CellContentClick += advancedDataGridView1_CellContentClick;
            advancedDataGridView1.CellDoubleClick += advancedDataGridView1_CellDoubleClick;
            // 
            // menuStrip1
            // 
            menuStrip1.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
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
            BackColor = Color.Gray;
            ClientSize = new Size(1902, 1033);
            Controls.Add(menuStrip1);
            Controls.Add(advancedDataGridView1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "UygunOlmayanListe";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "UygunOlmayanListe";
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
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fORMToolStripMenuItem;
        private ToolStripMenuItem eXCELÇEKToolStripMenuItem;
        private ToolStripMenuItem ePOSTAGÖNDERToolStripMenuItem;
        private ToolStripMenuItem lİSTEYİEXCELEAKTARToolStripMenuItem;
        private ToolStripMenuItem lİSTEToolStripMenuItem;
        private ToolStripMenuItem gERİDÖNToolStripMenuItem;
    }
}