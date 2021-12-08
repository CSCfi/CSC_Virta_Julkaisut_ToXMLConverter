namespace CSC_Virta_Julkaisut_ToXMLConverter
{
    partial class CSC_VIRTA_JulkaisutForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CSC_VIRTA_JulkaisutForm));
            this.LahdeDataGridView = new System.Windows.Forms.DataGridView();
            this.XMLdataGridView = new System.Windows.Forms.DataGridView();
            this.LueLahdeTiedostoButton = new System.Windows.Forms.Button();
            this.MuunnaXMLButton = new System.Windows.Forms.Button();
            this.TallennaXMLButton = new System.Windows.Forms.Button();
            this.ValidoiButton = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.customizeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.asetuksetToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.avaaLokitiedostoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.avaaCsvtarkistustiedostoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contentsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.Lahdelabel = new System.Windows.Forms.Label();
            this.XMLDatalabel = new System.Windows.Forms.Label();
            this.statusLabel = new System.Windows.Forms.Label();
            this.tiedostoNimiLabel = new System.Windows.Forms.Label();
            this.errorTextBox = new System.Windows.Forms.TextBox();
            this.AvaaXMLButton = new System.Windows.Forms.Button();
            this.Statuslabel2 = new System.Windows.Forms.Label();
            this.ValidoiXMLButton = new System.Windows.Forms.Button();
            this.mikaVuosilabel = new System.Windows.Forms.Label();
            this.numericUpDownVirhe = new System.Windows.Forms.NumericUpDown();
            this.numericUpDownRiviNro = new System.Windows.Forms.NumericUpDown();
            this.labelVirhenavigaattori = new System.Windows.Forms.Label();
            this.labelRiviNro = new System.Windows.Forms.Label();
            this.toolTipVirheNavigaattori = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.LahdeDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.XMLdataGridView)).BeginInit();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownVirhe)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownRiviNro)).BeginInit();
            this.SuspendLayout();
            // 
            // LahdeDataGridView
            // 
            this.LahdeDataGridView.AllowUserToAddRows = false;
            this.LahdeDataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.LahdeDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.LahdeDataGridView.Location = new System.Drawing.Point(13, 73);
            this.LahdeDataGridView.Name = "LahdeDataGridView";
            this.LahdeDataGridView.Size = new System.Drawing.Size(930, 189);
            this.LahdeDataGridView.TabIndex = 0;
            // 
            // XMLdataGridView
            // 
            this.XMLdataGridView.AllowUserToAddRows = false;
            this.XMLdataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.XMLdataGridView.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.XMLdataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.XMLdataGridView.Location = new System.Drawing.Point(13, 317);
            this.XMLdataGridView.Name = "XMLdataGridView";
            this.XMLdataGridView.Size = new System.Drawing.Size(930, 210);
            this.XMLdataGridView.TabIndex = 1;
            this.XMLdataGridView.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridVCellClick);
            // 
            // LueLahdeTiedostoButton
            // 
            this.LueLahdeTiedostoButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LueLahdeTiedostoButton.Location = new System.Drawing.Point(13, 271);
            this.LueLahdeTiedostoButton.Name = "LueLahdeTiedostoButton";
            this.LueLahdeTiedostoButton.Size = new System.Drawing.Size(200, 25);
            this.LueLahdeTiedostoButton.TabIndex = 2;
            this.LueLahdeTiedostoButton.Text = "Lue lähdetiedosto...";
            this.LueLahdeTiedostoButton.UseVisualStyleBackColor = true;
            this.LueLahdeTiedostoButton.Click += new System.EventHandler(this.LueLahdeTiedostoButton_Click);
            // 
            // MuunnaXMLButton
            // 
            this.MuunnaXMLButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MuunnaXMLButton.Location = new System.Drawing.Point(292, 271);
            this.MuunnaXMLButton.Name = "MuunnaXMLButton";
            this.MuunnaXMLButton.Size = new System.Drawing.Size(200, 25);
            this.MuunnaXMLButton.TabIndex = 3;
            this.MuunnaXMLButton.Text = "Muunna CSV lähdedata";
            this.MuunnaXMLButton.UseVisualStyleBackColor = true;
            this.MuunnaXMLButton.Click += new System.EventHandler(this.MuunnaXMLButton_Click);
            // 
            // TallennaXMLButton
            // 
            this.TallennaXMLButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TallennaXMLButton.Location = new System.Drawing.Point(13, 533);
            this.TallennaXMLButton.Name = "TallennaXMLButton";
            this.TallennaXMLButton.Size = new System.Drawing.Size(200, 25);
            this.TallennaXMLButton.TabIndex = 4;
            this.TallennaXMLButton.Tag = "Muodosta XML-tiedosto";
            this.TallennaXMLButton.Text = "Tallenna XML tiedostoon";
            this.TallennaXMLButton.UseVisualStyleBackColor = true;
            this.TallennaXMLButton.Click += new System.EventHandler(this.TallennaXMLButton_Click);
            // 
            // ValidoiButton
            // 
            this.ValidoiButton.FlatAppearance.BorderSize = 0;
            this.ValidoiButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ValidoiButton.Location = new System.Drawing.Point(557, 271);
            this.ValidoiButton.Name = "ValidoiButton";
            this.ValidoiButton.Size = new System.Drawing.Size(200, 25);
            this.ValidoiButton.TabIndex = 5;
            this.ValidoiButton.Tag = "Validoi aineisto";
            this.ValidoiButton.Text = "Validoi lähdedata";
            this.ValidoiButton.UseVisualStyleBackColor = true;
            this.ValidoiButton.Click += new System.EventHandler(this.ValidoiButton_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.toolsToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(955, 24);
            this.menuStrip1.TabIndex = 6;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem,
            this.toolStripSeparator,
            this.saveToolStripMenuItem,
            this.toolStripSeparator2,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E)));
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(65, 20);
            this.fileToolStripMenuItem.Text = "Ti&edosto";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("openToolStripMenuItem.Image")));
            this.openToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.openToolStripMenuItem.Size = new System.Drawing.Size(379, 22);
            this.openToolStripMenuItem.Text = "&Lue lähdetiedosto...";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.LueLahdeTiedostoButton_Click);
            // 
            // tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem
            // 
            this.tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem.Name = "tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem";
            this.tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
            this.tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem.Size = new System.Drawing.Size(379, 22);
            this.tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem.Text = "Tyhje&nnä vanhat tiedot ja lue uusi lähdetiedosto...";
            this.tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem.Click += new System.EventHandler(this.tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem_Click);
            // 
            // toolStripSeparator
            // 
            this.toolStripSeparator.Name = "toolStripSeparator";
            this.toolStripSeparator.Size = new System.Drawing.Size(376, 6);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("saveToolStripMenuItem.Image")));
            this.saveToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(379, 22);
            this.saveToolStripMenuItem.Text = "Tallenna uu&si XML tiedosto";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.TallennaXMLButton_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(376, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.X)));
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(379, 22);
            this.exitToolStripMenuItem.Text = "Poistu";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // toolsToolStripMenuItem
            // 
            this.toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.customizeToolStripMenuItem,
            this.optionsToolStripMenuItem,
            this.asetuksetToolStripMenuItem,
            this.avaaLokitiedostoToolStripMenuItem,
            this.avaaCsvtarkistustiedostoToolStripMenuItem});
            this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            this.toolsToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.U)));
            this.toolsToolStripMenuItem.Size = new System.Drawing.Size(65, 20);
            this.toolsToolStripMenuItem.Text = "Työkal&ut";
            // 
            // customizeToolStripMenuItem
            // 
            this.customizeToolStripMenuItem.Name = "customizeToolStripMenuItem";
            this.customizeToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.M)));
            this.customizeToolStripMenuItem.Size = new System.Drawing.Size(252, 22);
            this.customizeToolStripMenuItem.Text = "&Muunna XML muotoon";
            this.customizeToolStripMenuItem.Click += new System.EventHandler(this.MuunnaXMLButton_Click);
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.D)));
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(252, 22);
            this.optionsToolStripMenuItem.Text = "Vali&doi";
            this.optionsToolStripMenuItem.Click += new System.EventHandler(this.optionsToolStripMenuItem_Click);
            // 
            // asetuksetToolStripMenuItem
            // 
            this.asetuksetToolStripMenuItem.Name = "asetuksetToolStripMenuItem";
            this.asetuksetToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.A)));
            this.asetuksetToolStripMenuItem.Size = new System.Drawing.Size(252, 22);
            this.asetuksetToolStripMenuItem.Text = "&Asetukset";
            this.asetuksetToolStripMenuItem.Click += new System.EventHandler(this.asetuksetToolStripMenuItem_Click);
            // 
            // avaaLokitiedostoToolStripMenuItem
            // 
            this.avaaLokitiedostoToolStripMenuItem.Name = "avaaLokitiedostoToolStripMenuItem";
            this.avaaLokitiedostoToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.L)));
            this.avaaLokitiedostoToolStripMenuItem.Size = new System.Drawing.Size(252, 22);
            this.avaaLokitiedostoToolStripMenuItem.Text = "Avaa &lokitiedosto";
            this.avaaLokitiedostoToolStripMenuItem.Click += new System.EventHandler(this.avaaLokitiedostoToolStripMenuItem_Click);
            // 
            // avaaCsvtarkistustiedostoToolStripMenuItem
            // 
            this.avaaCsvtarkistustiedostoToolStripMenuItem.Name = "avaaCsvtarkistustiedostoToolStripMenuItem";
            this.avaaCsvtarkistustiedostoToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E)));
            this.avaaCsvtarkistustiedostoToolStripMenuItem.Size = new System.Drawing.Size(252, 22);
            this.avaaCsvtarkistustiedostoToolStripMenuItem.Text = "Avaa csv-tarkistustiedosto";
            this.avaaCsvtarkistustiedostoToolStripMenuItem.Click += new System.EventHandler(this.avaaCsvtarkistustiedostoToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.contentsToolStripMenuItem,
            this.toolStripSeparator5,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E)));
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Ohj&e";
            // 
            // contentsToolStripMenuItem
            // 
            this.contentsToolStripMenuItem.Name = "contentsToolStripMenuItem";
            this.contentsToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.H)));
            this.contentsToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.contentsToolStripMenuItem.Text = "O&hjeet";
            this.contentsToolStripMenuItem.Click += new System.EventHandler(this.contentsToolStripMenuItem_Click);
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(149, 6);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.T)));
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.aboutToolStripMenuItem.Text = "&Tiedot";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // Lahdelabel
            // 
            this.Lahdelabel.AutoSize = true;
            this.Lahdelabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Lahdelabel.Location = new System.Drawing.Point(12, 55);
            this.Lahdelabel.Name = "Lahdelabel";
            this.Lahdelabel.Size = new System.Drawing.Size(101, 15);
            this.Lahdelabel.TabIndex = 7;
            this.Lahdelabel.Text = "CSV lähdedata";
            // 
            // XMLDatalabel
            // 
            this.XMLDatalabel.AutoSize = true;
            this.XMLDatalabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.XMLDatalabel.Location = new System.Drawing.Point(12, 299);
            this.XMLDatalabel.Name = "XMLDatalabel";
            this.XMLDatalabel.Size = new System.Drawing.Size(107, 15);
            this.XMLDatalabel.TabIndex = 8;
            this.XMLDatalabel.Text = "Validointi dataa";
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(234, 534);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(145, 13);
            this.statusLabel.TabIndex = 9;
            this.statusLabel.Text = "Aineistoa ei ole vielä validoitu";
            // 
            // tiedostoNimiLabel
            // 
            this.tiedostoNimiLabel.AutoSize = true;
            this.tiedostoNimiLabel.Location = new System.Drawing.Point(142, 34);
            this.tiedostoNimiLabel.Name = "tiedostoNimiLabel";
            this.tiedostoNimiLabel.Size = new System.Drawing.Size(0, 13);
            this.tiedostoNimiLabel.TabIndex = 10;
            // 
            // errorTextBox
            // 
            this.errorTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.errorTextBox.Location = new System.Drawing.Point(15, 564);
            this.errorTextBox.Multiline = true;
            this.errorTextBox.Name = "errorTextBox";
            this.errorTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.errorTextBox.Size = new System.Drawing.Size(928, 59);
            this.errorTextBox.TabIndex = 11;
            // 
            // AvaaXMLButton
            // 
            this.AvaaXMLButton.Location = new System.Drawing.Point(557, 534);
            this.AvaaXMLButton.Name = "AvaaXMLButton";
            this.AvaaXMLButton.Size = new System.Drawing.Size(200, 23);
            this.AvaaXMLButton.TabIndex = 12;
            this.AvaaXMLButton.Tag = "Avaa XML";
            this.AvaaXMLButton.Text = "Avaa XML tiedosto";
            this.AvaaXMLButton.UseVisualStyleBackColor = true;
            this.AvaaXMLButton.Click += new System.EventHandler(this.AvaaXMLButton_Click);
            // 
            // Statuslabel2
            // 
            this.Statuslabel2.AutoSize = true;
            this.Statuslabel2.Location = new System.Drawing.Point(234, 548);
            this.Statuslabel2.Name = "Statuslabel2";
            this.Statuslabel2.Size = new System.Drawing.Size(16, 13);
            this.Statuslabel2.TabIndex = 13;
            this.Statuslabel2.Text = "...";
            // 
            // ValidoiXMLButton
            // 
            this.ValidoiXMLButton.Location = new System.Drawing.Point(772, 534);
            this.ValidoiXMLButton.Name = "ValidoiXMLButton";
            this.ValidoiXMLButton.Size = new System.Drawing.Size(171, 23);
            this.ValidoiXMLButton.TabIndex = 14;
            this.ValidoiXMLButton.Tag = "Validoi XML uudestaan";
            this.ValidoiXMLButton.Text = "Validoi XML";
            this.ValidoiXMLButton.UseVisualStyleBackColor = true;
            this.ValidoiXMLButton.Click += new System.EventHandler(this.ValidoiXMLButton_Click);
            // 
            // mikaVuosilabel
            // 
            this.mikaVuosilabel.AutoSize = true;
            this.mikaVuosilabel.Location = new System.Drawing.Point(15, 33);
            this.mikaVuosilabel.Name = "mikaVuosilabel";
            this.mikaVuosilabel.Size = new System.Drawing.Size(32, 13);
            this.mikaVuosilabel.TabIndex = 15;
            this.mikaVuosilabel.Text = "vuosi";
            // 
            // numericUpDownVirhe
            // 
            this.numericUpDownVirhe.Location = new System.Drawing.Point(878, 271);
            this.numericUpDownVirhe.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.numericUpDownVirhe.Name = "numericUpDownVirhe";
            this.numericUpDownVirhe.Size = new System.Drawing.Size(65, 20);
            this.numericUpDownVirhe.TabIndex = 16;
            this.numericUpDownVirhe.Tag = "Etsi seuraava virhe";
            this.numericUpDownVirhe.ValueChanged += new System.EventHandler(this.numericUpDownVirhe_ValueChanged);
            // 
            // numericUpDownRiviNro
            // 
            this.numericUpDownRiviNro.Enabled = false;
            this.numericUpDownRiviNro.Location = new System.Drawing.Point(878, 294);
            this.numericUpDownRiviNro.Maximum = new decimal(new int[] {
            100000,
            0,
            0,
            0});
            this.numericUpDownRiviNro.Name = "numericUpDownRiviNro";
            this.numericUpDownRiviNro.Size = new System.Drawing.Size(65, 20);
            this.numericUpDownRiviNro.TabIndex = 17;
            // 
            // labelVirhenavigaattori
            // 
            this.labelVirhenavigaattori.AutoSize = true;
            this.labelVirhenavigaattori.Location = new System.Drawing.Point(783, 271);
            this.labelVirhenavigaattori.Name = "labelVirhenavigaattori";
            this.labelVirhenavigaattori.Size = new System.Drawing.Size(89, 13);
            this.labelVirhenavigaattori.TabIndex = 18;
            this.labelVirhenavigaattori.Text = "Virhenavigaattori:";
            // 
            // labelRiviNro
            // 
            this.labelRiviNro.AutoSize = true;
            this.labelRiviNro.Location = new System.Drawing.Point(829, 296);
            this.labelRiviNro.Name = "labelRiviNro";
            this.labelRiviNro.Size = new System.Drawing.Size(43, 13);
            this.labelRiviNro.TabIndex = 19;
            this.labelRiviNro.Text = "Rivinro:";
            // 
            // CSC_VIRTA_JulkaisutForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.ClientSize = new System.Drawing.Size(955, 635);
            this.Controls.Add(this.labelRiviNro);
            this.Controls.Add(this.labelVirhenavigaattori);
            this.Controls.Add(this.numericUpDownRiviNro);
            this.Controls.Add(this.numericUpDownVirhe);
            this.Controls.Add(this.mikaVuosilabel);
            this.Controls.Add(this.ValidoiXMLButton);
            this.Controls.Add(this.Statuslabel2);
            this.Controls.Add(this.AvaaXMLButton);
            this.Controls.Add(this.errorTextBox);
            this.Controls.Add(this.tiedostoNimiLabel);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.XMLDatalabel);
            this.Controls.Add(this.Lahdelabel);
            this.Controls.Add(this.ValidoiButton);
            this.Controls.Add(this.TallennaXMLButton);
            this.Controls.Add(this.MuunnaXMLButton);
            this.Controls.Add(this.LueLahdeTiedostoButton);
            this.Controls.Add(this.XMLdataGridView);
            this.Controls.Add(this.LahdeDataGridView);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "CSC_VIRTA_JulkaisutForm";
            this.Text = "CSC VIRTA Julkaisutietojen CSV-XML-työkalu (versio 1.3.4 10.11.2021)";
            this.Load += new System.EventHandler(this.CSC_VIRTA_JulkaisutForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.LahdeDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.XMLdataGridView)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownVirhe)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownRiviNro)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView LahdeDataGridView;
        private System.Windows.Forms.DataGridView XMLdataGridView;
        private System.Windows.Forms.Button LueLahdeTiedostoButton;
        private System.Windows.Forms.Button MuunnaXMLButton;
        private System.Windows.Forms.Button TallennaXMLButton;
        private System.Windows.Forms.Button ValidoiButton;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem customizeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem contentsToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.Label Lahdelabel;
        private System.Windows.Forms.Label XMLDatalabel;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.Label tiedostoNimiLabel;
        private System.Windows.Forms.TextBox errorTextBox;
        private System.Windows.Forms.Button AvaaXMLButton;
        private System.Windows.Forms.Label Statuslabel2;
        private System.Windows.Forms.Button ValidoiXMLButton;
        private System.Windows.Forms.ToolStripMenuItem asetuksetToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem avaaLokitiedostoToolStripMenuItem;
        private System.Windows.Forms.Label mikaVuosilabel;
        private System.Windows.Forms.ToolStripMenuItem avaaCsvtarkistustiedostoToolStripMenuItem;
        private System.Windows.Forms.NumericUpDown numericUpDownVirhe;
        private System.Windows.Forms.NumericUpDown numericUpDownRiviNro;
        private System.Windows.Forms.Label labelVirhenavigaattori;
        private System.Windows.Forms.Label labelRiviNro;
        private System.Windows.Forms.ToolTip toolTipVirheNavigaattori;
    }
}

