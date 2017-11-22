namespace CSC_Virta_Julkaisut_ToXMLConverter
{
    partial class TarkistusAineisto
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TarkistusAineisto));
            this.TarkistusDataGridView = new System.Windows.Forms.DataGridView();
            this.TarkistusExcelButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.TarkistusDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // TarkistusDataGridView
            // 
            this.TarkistusDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.TarkistusDataGridView.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.TarkistusDataGridView.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.TarkistusDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.TarkistusDataGridView.Dock = System.Windows.Forms.DockStyle.Top;
            this.TarkistusDataGridView.GridColor = System.Drawing.SystemColors.ControlLight;
            this.TarkistusDataGridView.Location = new System.Drawing.Point(0, 0);
            this.TarkistusDataGridView.Name = "TarkistusDataGridView";
            this.TarkistusDataGridView.Size = new System.Drawing.Size(875, 313);
            this.TarkistusDataGridView.TabIndex = 0;
            // 
            // TarkistusExcelButton
            // 
            this.TarkistusExcelButton.Location = new System.Drawing.Point(12, 340);
            this.TarkistusExcelButton.Name = "TarkistusExcelButton";
            this.TarkistusExcelButton.Size = new System.Drawing.Size(138, 27);
            this.TarkistusExcelButton.TabIndex = 1;
            this.TarkistusExcelButton.Text = "Avaa Excelissä";
            this.TarkistusExcelButton.UseVisualStyleBackColor = true;
            this.TarkistusExcelButton.Click += new System.EventHandler(this.TarkistusExcelButton_Click);
            // 
            // TarkistusAineisto
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(875, 408);
            this.Controls.Add(this.TarkistusExcelButton);
            this.Controls.Add(this.TarkistusDataGridView);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "TarkistusAineisto";
            this.Text = "TarkistusAineisto";
            ((System.ComponentModel.ISupportInitialize)(this.TarkistusDataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.DataGridView TarkistusDataGridView;
        private System.Windows.Forms.Button TarkistusExcelButton;
    }
}