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
    public partial class frmDuzenle : Form
    {
        public Form1 frm1;
        public frmDuzenle()
        {
            InitializeComponent();
            
        }
        public static string sAdi = "";
        public static string sNo = "";
        public static string sSinif = "";
        public static string sSeriNo = "";
        public static string sSorun = "";
        public static string sTarih = "";
        public static bool cAriza = false;
        public static bool cGaranti = false;

        private void frmDuzenle_Load(object sender, EventArgs e)
        {

            txtAdi.Text = sAdi;
            txtNo.Text = sNo;
            string[] parcala = sSinif.Split('-');
            comboBox1.Text = parcala[0];
            comboBox2.Text = parcala[1];
            txtSeriNo.Text = sSeriNo;
            txtSorun.Text = sSorun;
            txtTarih.Text = sTarih;
            cbAriza.Checked = cAriza;
            cbGaranti.Checked = cGaranti;
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

        private void txtTarih_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
        OleDbConnection bag = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=data.accdb");
        OleDbCommand kmt = new OleDbCommand();
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtNo.Text.Trim() != "" && txtSeriNo.Text.Trim() != "")
                {
                    bag.Open();
                    kmt.Connection = bag;
                    kmt.CommandText = "UPDATE ogrbir SET adSoyad='" + txtAdi.Text + "',sinifi='" + comboBox1.Text + "-" + comboBox2.Text + "',ogrNo='" + txtNo.Text + "',tabletSorun='" + txtSorun.Text + "',arizaDurum='" + (cbAriza.Checked ? -1 : 0) + "',garanti='" + (cbGaranti.Checked ? -1 : 0) + "',tarih='" + txtTarih.Text + "' WHERE ogrNo='" + txtNo.Text + "'";
                    kmt.ExecuteNonQuery();
                    bag.Close();
                    this.Close();
                    Form1.tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir", bag);
                    adtr.Fill(Form1.tablo);
                    MessageBox.Show("Güncelleme İşlemi Tamamlandı.");
                    frm1.listele();
                }
                else
                {
                    MessageBox.Show("Yıldızlı alanları doldurmanız gerekiyor!");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString() + "\tKayıtlı Öğrenci No Girişi");
                bag.Close();
            }
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
