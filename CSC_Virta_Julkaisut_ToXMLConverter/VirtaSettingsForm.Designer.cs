namespace CSC_Virta_Julkaisut_ToXMLConverter
{
    partial class VirtaSettingsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(VirtaSettingsForm));
            this.SettingsTallennaButton = new System.Windows.Forms.Button();
            this.AsetuksetGroupBox = new System.Windows.Forms.GroupBox();
            this.VuosiRadioButton2016 = new System.Windows.Forms.RadioButton();
            this.VuosiRadioButton2015 = new System.Windows.Forms.RadioButton();
            this.SettingsPanel = new System.Windows.Forms.Panel();
            this.VuosiRadioButton2017 = new System.Windows.Forms.RadioButton();
            this.AsetuksetGroupBox.SuspendLayout();
            this.SettingsPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // SettingsTallennaButton
            // 
            this.SettingsTallennaButton.Location = new System.Drawing.Point(185, 28);
            this.SettingsTallennaButton.Name = "SettingsTallennaButton";
            this.SettingsTallennaButton.Size = new System.Drawing.Size(87, 23);
            this.SettingsTallennaButton.TabIndex = 0;
            this.SettingsTallennaButton.Text = "Tallenna";
            this.SettingsTallennaButton.UseVisualStyleBackColor = true;
            this.SettingsTallennaButton.Click += new System.EventHandler(this.SettingsTallennaButton_Click);
            // 
            // AsetuksetGroupBox
            // 
            this.AsetuksetGroupBox.Controls.Add(this.VuosiRadioButton2017);
            this.AsetuksetGroupBox.Controls.Add(this.VuosiRadioButton2016);
            this.AsetuksetGroupBox.Controls.Add(this.VuosiRadioButton2015);
            this.AsetuksetGroupBox.Location = new System.Drawing.Point(3, 3);
            this.AsetuksetGroupBox.Name = "AsetuksetGroupBox";
            this.AsetuksetGroupBox.Size = new System.Drawing.Size(164, 112);
            this.AsetuksetGroupBox.TabIndex = 5;
            this.AsetuksetGroupBox.TabStop = false;
            this.AsetuksetGroupBox.Text = "Ilmoitusvuosi:";
            // 
            // VuosiRadioButton2016
            // 
            this.VuosiRadioButton2016.AutoSize = true;
            this.VuosiRadioButton2016.Checked = true;
            this.VuosiRadioButton2016.Location = new System.Drawing.Point(84, 42);
            this.VuosiRadioButton2016.Name = "VuosiRadioButton2016";
            this.VuosiRadioButton2016.Size = new System.Drawing.Size(49, 17);
            this.VuosiRadioButton2016.TabIndex = 1;
            this.VuosiRadioButton2016.TabStop = true;
            this.VuosiRadioButton2016.Text = "2016";
            this.VuosiRadioButton2016.UseVisualStyleBackColor = true;
            // 
            // VuosiRadioButton2015
            // 
            this.VuosiRadioButton2015.AutoSize = true;
            this.VuosiRadioButton2015.Location = new System.Drawing.Point(84, 19);
            this.VuosiRadioButton2015.Name = "VuosiRadioButton2015";
            this.VuosiRadioButton2015.Size = new System.Drawing.Size(49, 17);
            this.VuosiRadioButton2015.TabIndex = 0;
            this.VuosiRadioButton2015.Text = "2015";
            this.VuosiRadioButton2015.UseVisualStyleBackColor = true;
            // 
            // SettingsPanel
            // 
            this.SettingsPanel.Controls.Add(this.AsetuksetGroupBox);
            this.SettingsPanel.Location = new System.Drawing.Point(12, 12);
            this.SettingsPanel.Name = "SettingsPanel";
            this.SettingsPanel.Size = new System.Drawing.Size(170, 238);
            this.SettingsPanel.TabIndex = 1;
            // 
            // VuosiRadioButton2017
            // 
            this.VuosiRadioButton2017.AutoSize = true;
            this.VuosiRadioButton2017.Location = new System.Drawing.Point(84, 65);
            this.VuosiRadioButton2017.Name = "VuosiRadioButton2017";
            this.VuosiRadioButton2017.Size = new System.Drawing.Size(49, 17);
            this.VuosiRadioButton2017.TabIndex = 2;
            this.VuosiRadioButton2017.Text = "2017";
            this.VuosiRadioButton2017.UseVisualStyleBackColor = true;
            // 
            // VirtaSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.SettingsPanel);
            this.Controls.Add(this.SettingsTallennaButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "VirtaSettingsForm";
            this.Text = "Asetukset";
            this.AsetuksetGroupBox.ResumeLayout(false);
            this.AsetuksetGroupBox.PerformLayout();
            this.SettingsPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button SettingsTallennaButton;
        private System.Windows.Forms.GroupBox AsetuksetGroupBox;
        private System.Windows.Forms.RadioButton VuosiRadioButton2016;
        private System.Windows.Forms.RadioButton VuosiRadioButton2015;
        private System.Windows.Forms.Panel SettingsPanel;
        private System.Windows.Forms.RadioButton VuosiRadioButton2017;
    }
}