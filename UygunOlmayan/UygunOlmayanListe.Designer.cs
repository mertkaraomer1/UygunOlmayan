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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UygunOlmayanListe));
            advancedDataGridView1 = new Zuby.ADGV.AdvancedDataGridView();
            toolStrip1 = new ToolStrip();
            toolStripLabel1 = new ToolStripLabel();
            toolStripButton1 = new ToolStripButton();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).BeginInit();
            toolStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // advancedDataGridView1
            // 
            advancedDataGridView1.AllowUserToAddRows = false;
            advancedDataGridView1.AllowUserToDeleteRows = false;
            advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            advancedDataGridView1.BackgroundColor = Color.White;
            advancedDataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            advancedDataGridView1.FilterAndSortEnabled = true;
            advancedDataGridView1.FilterStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.Location = new Point(24, 71);
            advancedDataGridView1.Name = "advancedDataGridView1";
            advancedDataGridView1.RightToLeft = RightToLeft.No;
            advancedDataGridView1.RowHeadersWidth = 51;
            advancedDataGridView1.RowTemplate.Height = 29;
            advancedDataGridView1.Size = new Size(1843, 925);
            advancedDataGridView1.SortStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.TabIndex = 3;
            advancedDataGridView1.CellClick += advancedDataGridView1_CellClick;
            advancedDataGridView1.CellContentClick += advancedDataGridView1_CellContentClick;
            advancedDataGridView1.CellDoubleClick += advancedDataGridView1_CellDoubleClick;
            // 
            // toolStrip1
            // 
            toolStrip1.BackColor = Color.DimGray;
            toolStrip1.ImageScalingSize = new Size(20, 20);
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripLabel1, toolStripButton1 });
            toolStrip1.Location = new Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(1902, 27);
            toolStrip1.TabIndex = 20;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel1
            // 
            toolStripLabel1.Name = "toolStripLabel1";
            toolStripLabel1.Size = new Size(0, 24);
            // 
            // toolStripButton1
            // 
            toolStripButton1.DisplayStyle = ToolStripItemDisplayStyle.Image;
            toolStripButton1.Image = (Image)resources.GetObject("toolStripButton1.Image");
            toolStripButton1.ImageTransparentColor = Color.Magenta;
            toolStripButton1.Name = "toolStripButton1";
            toolStripButton1.Size = new Size(29, 24);
            toolStripButton1.Text = "toolStripButton1";
            toolStripButton1.Click += toolStripButton1_Click;
            // 
            // UygunOlmayanListe
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Gray;
            ClientSize = new Size(1902, 1033);
            Controls.Add(toolStrip1);
            Controls.Add(advancedDataGridView1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "UygunOlmayanListe";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "UygunOlmayanListe";
            WindowState = FormWindowState.Maximized;
            FormClosing += UygunOlmayanListe_FormClosing;
            Load += UygunOlmayanListe_Load;
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).EndInit();
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private Zuby.ADGV.AdvancedDataGridView advancedDataGridView1;
        private ToolStrip toolStrip1;
        private ToolStripLabel toolStripLabel1;
        private ToolStripButton toolStripButton1;
    }
}