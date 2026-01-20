namespace UygunOlmayan
{
    partial class UygunOlmayanRapor
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
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle3 = new DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UygunOlmayanRapor));
            toolStrip1 = new ToolStrip();
            toolStripLabel1 = new ToolStripLabel();
            toolStripButton1 = new ToolStripButton();
            toolStripButton2 = new ToolStripButton();
            advancedDataGridView1 = new Zuby.ADGV.AdvancedDataGridView();
            panelHeader = new Panel();
            lblTitle = new Label();
            toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).BeginInit();
            SuspendLayout();
            // 
            // toolStrip1
            // 
            toolStrip1.BackColor = Color.FromArgb(41, 57, 85);
            toolStrip1.GripStyle = ToolStripGripStyle.Hidden;
            toolStrip1.ImageScalingSize = new Size(20, 20);
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripLabel1, toolStripButton1, toolStripButton2 });
            toolStrip1.Location = new Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(1902, 35);
            toolStrip1.TabIndex = 22;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel1
            // 
            toolStripLabel1.Name = "toolStripLabel1";
            toolStripLabel1.Size = new Size(0, 24);
            // 
            // toolStripButton1
            // 
            toolStripButton1.ForeColor = Color.White;
            toolStripButton1.Image = (Image)resources.GetObject("toolStripButton1.Image");
            toolStripButton1.ImageTransparentColor = Color.Magenta;
            toolStripButton1.Name = "toolStripButton1";
            toolStripButton1.Size = new Size(80, 32);
            toolStripButton1.Text = "YENİLE";
            toolStripButton1.Click += toolStripButton1_Click;
            // 
            // toolStripButton2
            // 
            toolStripButton2.ForeColor = Color.White;
            toolStripButton2.Image = (Image)resources.GetObject("toolStripButton2.Image");
            toolStripButton2.ImageTransparentColor = Color.Magenta;
            toolStripButton2.Name = "toolStripButton2";
            toolStripButton2.Size = new Size(100, 32);
            toolStripButton2.Text = "EXCEL'E AKTAR";
            toolStripButton2.Click += toolStripButton2_Click;
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
            advancedDataGridView1.Location = new Point(12, 130);
            advancedDataGridView1.Name = "advancedDataGridView1";
            advancedDataGridView1.ReadOnly = true;
            advancedDataGridView1.RowHeadersVisible = false;
            advancedDataGridView1.RowHeadersWidth = 51;
            dataGridViewCellStyle3.BackColor = Color.FromArgb(250, 251, 252);
            advancedDataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            advancedDataGridView1.RowTemplate.Height = 40;
            advancedDataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            advancedDataGridView1.Size = new Size(1878, 891);
            advancedDataGridView1.SortStringChangedInvokeBeforeDatasourceUpdate = true;
            advancedDataGridView1.TabIndex = 21;
            // 
            // panelHeader
            // 
            panelHeader.BackColor = Color.FromArgb(41, 57, 85);
            panelHeader.Controls.Add(lblTitle);
            panelHeader.Dock = DockStyle.Top;
            panelHeader.Location = new Point(0, 35);
            panelHeader.Name = "panelHeader";
            panelHeader.Size = new Size(1902, 80);
            panelHeader.TabIndex = 23;
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
            lblTitle.Text = "UYGUN OLMAYAN RAPOR LİSTESİ";
            // 
            // UygunOlmayanRapor
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(242, 245, 250);
            ClientSize = new Size(1902, 1033);
            Controls.Add(panelHeader);
            Controls.Add(toolStrip1);
            Controls.Add(advancedDataGridView1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "UygunOlmayanRapor";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Uygun Olmayan Rapor";
            WindowState = FormWindowState.Maximized;
            Load += UygunOlmayanRapor_Load;
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)advancedDataGridView1).EndInit();
            panelHeader.ResumeLayout(false);
            panelHeader.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ToolStrip toolStrip1;
        private ToolStripLabel toolStripLabel1;
        private ToolStripButton toolStripButton1;
        private Zuby.ADGV.AdvancedDataGridView advancedDataGridView1;
        private ToolStripButton toolStripButton2;
        private Panel panelHeader;
        private Label lblTitle;
    }
}