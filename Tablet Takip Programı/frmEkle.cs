using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Tablet_Takip_Programı
{
    public partial class frmEkle : Form
    {
        public Form1 frm1;
        public frmEkle()
        {
            InitializeComponent();
            frm1 = new Form1();
        }

        private void txtAdi_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar) == false && e.KeyChar != (char)08 && char.IsSeparator(e.KeyChar) == false)
            {
                e.Handled = true;
            }
        }

        private void txtNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsDigit(e.KeyChar) == false && e.KeyChar != (char)08)
            {
                e.Handled = true;
            }
        }

        private void txtNo_TextChanged(object sender, EventArgs e)
        {
            if (txtNo.Text != "")
            {
                pictureBox1.BackgroundImage = Tablet_Takip_Programı.Properties.Resources.tick1;
                pictureBox1.Visible = true;
            }

            else
            {
                pictureBox1.BackgroundImage = Tablet_Takip_Programı.Properties.Resources.warning1;
                pictureBox1.Visible = true;
            }
        }

        private void txtSeriNo_TextChanged(object sender, EventArgs e)
        {
            if (txtSeriNo.Text != "")
            {
                pictureBox2.BackgroundImage = Tablet_Takip_Programı.Properties.Resources.tick1;
                pictureBox2.Visible = true;
            }

            else
            {
                pictureBox2.BackgroundImage = Tablet_Takip_Programı.Properties.Resources.warning1;
                pictureBox2.Visible = true;
            }
        }
        OleDbConnection bag = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=data.accdb");
        OleDbCommand kmt = new OleDbCommand();
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                bag.Open();
                kmt.Connection = bag;
                kmt.CommandText = "INSERT INTO ogrbir (adSoyad,sinifi,ogrNo,tabletSeriNo,arizaDurum,garanti,tabletSorun,tarih) VALUES ('" + (txtAdi.Text) + "','" + (comboBox1.Text + "-" + comboBox2.Text) + "','" + (txtNo.Text) + "','" + (txtSeriNo.Text) + "'," + (cbAriza.Checked) + "," + (cbGaranti.Checked) + ",'" + (txtSorun.Text) + "','" + DateTime.Now + "')";
                kmt.ExecuteNonQuery();
                bag.Close();
                frm1.listele();
                this.Close();
                Form1.tablo.Clear();
                OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir", bag);
                adtr.Fill(Form1.tablo);
                MessageBox.Show("Kayıt İşlemi Gerçekleşti !");
            }
            catch (Exception ex)
            {
                bag.Close();
                MessageBox.Show(ex.ToString());
            }
        }

        private void frmEkle_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "9";
            comboBox2.Text = "A";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
