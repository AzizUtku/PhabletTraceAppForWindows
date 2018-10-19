namespace Tablet_Takip_Programı
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnSil = new System.Windows.Forms.Button();
            this.btnDuzenle = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.txtSort = new System.Windows.Forms.TextBox();
            this.cmbSortAriza = new System.Windows.Forms.ComboBox();
            this.cmbSortGaranti = new System.Windows.Forms.ComboBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.dosyaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exceleÇıkartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.yenileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.çıkışToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.öğrenciİşlemleriToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.yeniÖğrenciEkleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.öğrenciDüzenleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.öğrenciSilToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.hakkındaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button4 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(10, 26);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(171, 39);
            this.button1.TabIndex = 0;
            this.button1.Text = "Yeni Öğrenci Ekle";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 240);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(1036, 384);
            this.dataGridView1.TabIndex = 22;
            // 
            // btnSil
            // 
            this.btnSil.Location = new System.Drawing.Point(366, 26);
            this.btnSil.Margin = new System.Windows.Forms.Padding(4);
            this.btnSil.Name = "btnSil";
            this.btnSil.Size = new System.Drawing.Size(112, 39);
            this.btnSil.TabIndex = 2;
            this.btnSil.Text = "Öğrenci Sil";
            this.btnSil.UseVisualStyleBackColor = true;
            this.btnSil.Click += new System.EventHandler(this.btnSil_Click);
            // 
            // btnDuzenle
            // 
            this.btnDuzenle.Location = new System.Drawing.Point(188, 26);
            this.btnDuzenle.Name = "btnDuzenle";
            this.btnDuzenle.Size = new System.Drawing.Size(171, 39);
            this.btnDuzenle.TabIndex = 1;
            this.btnDuzenle.Text = "Öğrenci Düzenle";
            this.btnDuzenle.UseVisualStyleBackColor = true;
            this.btnDuzenle.Click += new System.EventHandler(this.btnDuzenle_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(486, 26);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(185, 39);
            this.button2.TabIndex = 3;
            this.button2.Text = "Excel \'e Çıkart";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(679, 26);
            this.button3.Margin = new System.Windows.Forms.Padding(4);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(77, 39);
            this.button3.TabIndex = 4;
            this.button3.Text = "Yenile";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // printDocument1
            // 
            this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.ItemHeight = 20;
            this.comboBox1.Items.AddRange(new object[] {
            "İsme Göre",
            "Sınıfa Göre",
            "No\'ya Göre",
            "Arızaya Göre",
            "Garantiye Göre",
            "Seri No\'ya Göre",
            "Soruna Göre",
            "Tarihe Göre"});
            this.comboBox1.Location = new System.Drawing.Point(6, 25);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(166, 28);
            this.comboBox1.TabIndex = 42;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // txtSort
            // 
            this.txtSort.Location = new System.Drawing.Point(178, 25);
            this.txtSort.Name = "txtSort";
            this.txtSort.Size = new System.Drawing.Size(99, 26);
            this.txtSort.TabIndex = 43;
            this.txtSort.TextChanged += new System.EventHandler(this.txtSort_TextChanged);
            // 
            // cmbSortAriza
            // 
            this.cmbSortAriza.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSortAriza.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbSortAriza.FormattingEnabled = true;
            this.cmbSortAriza.Items.AddRange(new object[] {
            "Arızalı",
            "Sağlam",
            "Sıfırla"});
            this.cmbSortAriza.Location = new System.Drawing.Point(178, 25);
            this.cmbSortAriza.Name = "cmbSortAriza";
            this.cmbSortAriza.Size = new System.Drawing.Size(99, 28);
            this.cmbSortAriza.TabIndex = 44;
            this.cmbSortAriza.Visible = false;
            // 
            // cmbSortGaranti
            // 
            this.cmbSortGaranti.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSortGaranti.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbSortGaranti.FormattingEnabled = true;
            this.cmbSortGaranti.ItemHeight = 20;
            this.cmbSortGaranti.Items.AddRange(new object[] {
            "Evet",
            "Hayır",
            "Sıfırla"});
            this.cmbSortGaranti.Location = new System.Drawing.Point(178, 25);
            this.cmbSortGaranti.Name = "cmbSortGaranti";
            this.cmbSortGaranti.Size = new System.Drawing.Size(99, 28);
            this.cmbSortGaranti.TabIndex = 46;
            this.cmbSortGaranti.Visible = false;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.dosyaToolStripMenuItem,
            this.öğrenciİşlemleriToolStripMenuItem,
            this.hakkındaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1060, 24);
            this.menuStrip1.TabIndex = 47;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // dosyaToolStripMenuItem
            // 
            this.dosyaToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exceleÇıkartToolStripMenuItem,
            this.yenileToolStripMenuItem,
            this.çıkışToolStripMenuItem});
            this.dosyaToolStripMenuItem.Name = "dosyaToolStripMenuItem";
            this.dosyaToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.dosyaToolStripMenuItem.Text = "Menü";
            this.dosyaToolStripMenuItem.Click += new System.EventHandler(this.dosyaToolStripMenuItem_Click);
            // 
            // exceleÇıkartToolStripMenuItem
            // 
            this.exceleÇıkartToolStripMenuItem.Name = "exceleÇıkartToolStripMenuItem";
            this.exceleÇıkartToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.E)));
            this.exceleÇıkartToolStripMenuItem.Size = new System.Drawing.Size(218, 22);
            this.exceleÇıkartToolStripMenuItem.Text = "Excel \'e Çıkart";
            this.exceleÇıkartToolStripMenuItem.Click += new System.EventHandler(this.exceleÇıkartToolStripMenuItem_Click);
            // 
            // yenileToolStripMenuItem
            // 
            this.yenileToolStripMenuItem.Name = "yenileToolStripMenuItem";
            this.yenileToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F5;
            this.yenileToolStripMenuItem.Size = new System.Drawing.Size(218, 22);
            this.yenileToolStripMenuItem.Text = "Yenile";
            this.yenileToolStripMenuItem.Click += new System.EventHandler(this.yenileToolStripMenuItem_Click);
            // 
            // çıkışToolStripMenuItem
            // 
            this.çıkışToolStripMenuItem.Name = "çıkışToolStripMenuItem";
            this.çıkışToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4)));
            this.çıkışToolStripMenuItem.Size = new System.Drawing.Size(218, 22);
            this.çıkışToolStripMenuItem.Text = "Çıkış";
            this.çıkışToolStripMenuItem.Click += new System.EventHandler(this.çıkışToolStripMenuItem_Click);
            // 
            // öğrenciİşlemleriToolStripMenuItem
            // 
            this.öğrenciİşlemleriToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.yeniÖğrenciEkleToolStripMenuItem,
            this.öğrenciDüzenleToolStripMenuItem,
            this.öğrenciSilToolStripMenuItem});
            this.öğrenciİşlemleriToolStripMenuItem.Name = "öğrenciİşlemleriToolStripMenuItem";
            this.öğrenciİşlemleriToolStripMenuItem.Size = new System.Drawing.Size(108, 20);
            this.öğrenciİşlemleriToolStripMenuItem.Text = "Öğrenci İşlemleri";
            // 
            // yeniÖğrenciEkleToolStripMenuItem
            // 
            this.yeniÖğrenciEkleToolStripMenuItem.Name = "yeniÖğrenciEkleToolStripMenuItem";
            this.yeniÖğrenciEkleToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E)));
            this.yeniÖğrenciEkleToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.yeniÖğrenciEkleToolStripMenuItem.Text = "Yeni Öğrenci Ekle";
            this.yeniÖğrenciEkleToolStripMenuItem.Click += new System.EventHandler(this.yeniÖğrenciEkleToolStripMenuItem_Click);
            // 
            // öğrenciDüzenleToolStripMenuItem
            // 
            this.öğrenciDüzenleToolStripMenuItem.Name = "öğrenciDüzenleToolStripMenuItem";
            this.öğrenciDüzenleToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.D)));
            this.öğrenciDüzenleToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.öğrenciDüzenleToolStripMenuItem.Text = "Öğrenci Düzenle";
            this.öğrenciDüzenleToolStripMenuItem.Click += new System.EventHandler(this.öğrenciDüzenleToolStripMenuItem_Click);
            // 
            // öğrenciSilToolStripMenuItem
            // 
            this.öğrenciSilToolStripMenuItem.Name = "öğrenciSilToolStripMenuItem";
            this.öğrenciSilToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.S)));
            this.öğrenciSilToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.öğrenciSilToolStripMenuItem.Text = "Öğrenci Sil";
            this.öğrenciSilToolStripMenuItem.Click += new System.EventHandler(this.öğrenciSilToolStripMenuItem_Click);
            // 
            // hakkındaToolStripMenuItem
            // 
            this.hakkındaToolStripMenuItem.Name = "hakkındaToolStripMenuItem";
            this.hakkındaToolStripMenuItem.Size = new System.Drawing.Size(69, 20);
            this.hakkındaToolStripMenuItem.Text = "Hakkında";
            this.hakkındaToolStripMenuItem.Click += new System.EventHandler(this.hakkındaToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Firebrick;
            this.label1.Location = new System.Drawing.Point(12, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(1036, 34);
            this.label1.TabIndex = 48;
            this.label1.Text = "Aziz Utku Kağıtcı Hoşgeldiniz!";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.cmbSortGaranti);
            this.groupBox1.Controls.Add(this.pictureBox1);
            this.groupBox1.Controls.Add(this.cmbSortAriza);
            this.groupBox1.Controls.Add(this.txtSort);
            this.groupBox1.Location = new System.Drawing.Point(376, 81);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(325, 64);
            this.groupBox1.TabIndex = 49;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Arama";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = global::Tablet_Takip_Programı.Properties.Resources.search;
            this.pictureBox1.Cursor = System.Windows.Forms.Cursors.Default;
            this.pictureBox1.Location = new System.Drawing.Point(283, 23);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(32, 32);
            this.pictureBox1.TabIndex = 45;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            this.pictureBox1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pictureBox1_MouseDown);
            this.pictureBox1.MouseLeave += new System.EventHandler(this.pictureBox1_MouseLeave);
            this.pictureBox1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.pictureBox1_MouseMove);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Controls.Add(this.btnSil);
            this.groupBox2.Controls.Add(this.btnDuzenle);
            this.groupBox2.Controls.Add(this.button3);
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Location = new System.Drawing.Point(142, 151);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(769, 73);
            this.groupBox2.TabIndex = 50;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "İşlemler";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(82, 81);
            this.button4.Margin = new System.Windows.Forms.Padding(4);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(185, 39);
            this.button4.TabIndex = 5;
            this.button4.Text = "Excel \'den Ekle";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Visible = false;
            this.button4.Click += new System.EventHandler(this.button4_Click_1);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(1060, 641);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Miramonte", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Tablet Takip Programı";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ToolTip toolTip1;
        public System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnSil;
        private System.Windows.Forms.Button btnDuzenle;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.TextBox txtSort;
        private System.Windows.Forms.ComboBox cmbSortAriza;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ComboBox cmbSortGaranti;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem dosyaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exceleÇıkartToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem yenileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem öğrenciİşlemleriToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem hakkındaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem yeniÖğrenciEkleToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem öğrenciDüzenleToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem öğrenciSilToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ToolStripMenuItem çıkışToolStripMenuItem;
        private System.Windows.Forms.Button button4;
    }
}

