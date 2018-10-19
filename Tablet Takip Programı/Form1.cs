using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Tablet_Takip_Programı
{
    public partial class Form1 : Form
    {
        public frmDuzenle frmDuzen;
        public frmEkle frmEk;
        public Form1()
        {
            InitializeComponent();
        }
        public static string kullaniciadi = "" ;
        OleDbConnection bag = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=data.accdb");
        OleDbCommand kmt = new OleDbCommand();
        private void button1_Click(object sender, EventArgs e)
        {
            frmEk = new frmEkle();
            frmEk.ShowDialog();
        }
        public static DataTable tablo = new DataTable();
        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "İsme Göre";
            cmbSortAriza.Text = "Arızalı";
            cmbSortGaranti.Text = "Evet";
            if(kullaniciadi == "azizutku"){
                label1.Text = "Aziz Utku Kağıtcı Hoşgeldiniz!";
            }
            else if (kullaniciadi == "nesehayta")
            {
                label1.Text = "Neşe Hayta Hoşgeldiniz!";
            }
            else if (kullaniciadi == "misafir")
            {
                label1.Text = "Misafir Hoşgeldiniz!";
            } else {
                label1.Text = kullaniciadi +" Hoşgeldiniz!";
            }
            dataGridView1.DataSource = tablo;
            listele();
            dataGridView1.Columns[0].HeaderText = "Adı Soyadı";
            dataGridView1.Columns[0].Width = 140;
            dataGridView1.Columns[1].HeaderText = "Sınıfı";
            dataGridView1.Columns[1].Width = 45;
            dataGridView1.Columns[2].HeaderText = "Öğrenci No";
            dataGridView1.Columns[2].Width = 65;
            dataGridView1.Columns[3].HeaderText = "Tablet Seri No";
            dataGridView1.Columns[3].Width = 140;
            dataGridView1.Columns[4].HeaderText = "Arıza Durumu";
            dataGridView1.Columns[4].Width = 65;
            dataGridView1.Columns[5].HeaderText = "Garantide";
            dataGridView1.Columns[5].Width = 87;
            dataGridView1.Columns[6].HeaderText = "Tabletin Sorunu";
            dataGridView1.Columns[6].Width = 311;
            dataGridView1.Columns[7].HeaderText = "Tarih";
            dataGridView1.Columns[7].Width = 140;
          //  dataGridView1.Columns[6].Width = 30;
        }

        public void listele()
        {
            try
            {
                tablo.Clear();
                OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir", bag);
                adtr.Fill(tablo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
           
        }

        private void txtNo_TextChanged(object sender, EventArgs e)
        {
         
        }

       

        private void txtNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsDigit(e.KeyChar) == false && e.KeyChar != (char)08)
            {
                e.Handled = true;
            }
        }

        private void txtAdi_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar) == false && e.KeyChar != (char)08 && char.IsSeparator(e.KeyChar) == false)
            {
                e.Handled = true;
            }
        }

        private void btnSil_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult cevap;
                cevap = MessageBox.Show("Kaydı silmek istediğinizden emin misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes)
                {
                    bag.Open();
                    kmt.Connection = bag;
                    kmt.CommandText = "DELETE FROM ogrbir WHERE ogrNo='" + dataGridView1.CurrentRow.Cells[2].Value.ToString() + "'";
                    kmt.ExecuteNonQuery();
                    bag.Close();
                    listele();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                bag.Close();
            }
            
        }

   
        private void btnDuzenle_Click(object sender, EventArgs e)
        {
            frmDuzen = new frmDuzenle();
            frmDuzenle.sAdi = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            frmDuzenle.sSinif = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            frmDuzenle.sNo = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            frmDuzenle.sSeriNo = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            frmDuzenle.cAriza = (bool)dataGridView1.CurrentRow.Cells[4].Value;
            frmDuzenle.cGaranti = (bool)dataGridView1.CurrentRow.Cells[5].Value;
            frmDuzenle.sSorun = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            frmDuzenle.sTarih = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            frmDuzen.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
          
            Excel.Application excelAktar = new Excel.Application();
            excelAktar.Visible = true;
            Excel.Workbook workbook = excelAktar.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet sheet1 = (Excel.Worksheet)excelAktar.ActiveSheet;
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Excel.Range myRange = (Excel.Range)sheet1.Cells[1, j + 1];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                myRange.ColumnWidth = 18;
             //   myRange.Columns.Font.Color = ColorTranslator.ToOle(Color.Red) ;
                myRange.Columns.Font.Bold = true;
                myRange.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            StartRow++;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    try
                    {
                        Excel.Range myRange = (Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];

                        string ss = dataGridView1[j, i].Value.ToString();
                        object sss = dataGridView1[j, i].Value;
                        myRange.Value2 = dataGridView1[j, i].Value == null || dataGridView1[j, i].Value is bool ? (j == 5 ? (bool)dataGridView1[j, i].Value == false ? "Hayır" : "Evet" : (bool)dataGridView1[j, i].Value == true ? "Arızalı" : "Sağlam") : (dataGridView1[j, i].Value);
                        myRange.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                    }
                    catch(Exception ex)
                    {
                        if (kullaniciadi == "azizutku")
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                }
            }
            
        }


       
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listele();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            printDocument1.Print();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap bm = new Bitmap(dataGridView1.Width, dataGridView1.Height);
            dataGridView1.DrawToBitmap(bm, new Rectangle(0, 0, dataGridView1.Width, dataGridView1.Height));
            e.Graphics.DrawImage(bm, 0, 0);
        }

        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            pictureBox1.BackgroundImage = Tablet_Takip_Programı.Properties.Resources.search_hover;
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_MouseLeave(object sender, EventArgs e)
        {
            pictureBox1.BackgroundImage = Tablet_Takip_Programı.Properties.Resources.search;
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox1.BackgroundImage = Tablet_Takip_Programı.Properties.Resources.search_hover2;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.Text=="Arızaya Göre"){
                txtSort.Visible = false;
                cmbSortGaranti.Visible = false;
                cmbSortAriza.Visible = true;
            }

            else if (comboBox1.Text == "Garantiye Göre")
            {
                txtSort.Visible = false;
                cmbSortGaranti.Visible = true;
                cmbSortAriza.Visible = false;
            }

            else
            {
                txtSort.Visible = true;
                cmbSortGaranti.Visible = false;
                cmbSortAriza.Visible = false;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Arızaya Göre")
            {
                if (cmbSortAriza.Text == "Arızalı")
                {
                    tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where arizaDurum Like '%" + -1 + "%'", bag);
                    adtr.Fill(tablo);
                    dataGridView1.DataSource = tablo;
                }

                else if(cmbSortAriza.Text == "Sağlam")
                {
                    tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where arizaDurum Like '%" + 0 + "%'", bag);
                    adtr.Fill(tablo);
                    dataGridView1.DataSource = tablo;
                }

                else
                {
                        tablo.Clear();
                        OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir", bag);
                        adtr.Fill(tablo);
                        dataGridView1.DataSource = tablo;
                }
            }

            else if (comboBox1.Text == "Garantiye Göre")
            {
                if (cmbSortGaranti.Text == "Evet")
                {
                    tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where garanti Like '%" + -1 + "%'", bag);
                    adtr.Fill(tablo);
                    dataGridView1.DataSource = tablo;
                }

                else if (cmbSortGaranti.Text == "Hayır")
                {
                    tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where garanti Like '%" + 0 + "%'", bag);
                    adtr.Fill(tablo);
                    dataGridView1.DataSource = tablo;
                }

                else
                {
                        tablo.Clear();
                        OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir", bag);
                        adtr.Fill(tablo);
                        dataGridView1.DataSource = tablo;
                   
                }

            }

            else if(comboBox1.Text == "İsme Göre")
            {
                    tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where adSoyad Like '%" + txtSort.Text + "%'", bag);
                    adtr.Fill(tablo);
                    dataGridView1.DataSource = tablo;
                
            }

            else if (comboBox1.Text == "Sınıfa Göre")
            {

                tablo.Clear();
                OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where sinifi Like '%" + txtSort.Text + "%'", bag);
                adtr.Fill(tablo);
                dataGridView1.DataSource = tablo;

            }

            else if (comboBox1.Text == "No'ya Göre")
            {

                tablo.Clear();
                OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where ogrNo Like '%" + txtSort.Text + "%'", bag);
                adtr.Fill(tablo);
                dataGridView1.DataSource = tablo;

            }

            else if (comboBox1.Text == "Seri No'ya Göre")
            {

                tablo.Clear();
                OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where tabletSeriNo Like '%" + txtSort.Text + "%'", bag);
                adtr.Fill(tablo);
                dataGridView1.DataSource = tablo;

            }


            else if (comboBox1.Text == "Soruna Göre")
            {

                tablo.Clear();
                OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where tabletSorun Like '%" + txtSort.Text + "%'", bag);
                adtr.Fill(tablo);
                dataGridView1.DataSource = tablo;

            }

            else if (comboBox1.Text == "Tarihe Göre")
            {

                tablo.Clear();
                OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where tarih Like '%" + txtSort.Text + "%'", bag);
                adtr.Fill(tablo);
                dataGridView1.DataSource = tablo;

            }
        }

        private void txtSort_TextChanged(object sender, EventArgs e)
        {
            if (txtSort.Text.Trim() == "")
            {
                tablo.Clear();
                OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir", bag);
                adtr.Fill(tablo);
                dataGridView1.DataSource = tablo;
            }

            else {

                if (comboBox1.Text == "İsme Göre")
                {
                    tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where adSoyad Like '%" + txtSort.Text + "%'", bag);
                    adtr.Fill(tablo);
                    dataGridView1.DataSource = tablo;

                }

                else if (comboBox1.Text == "Sınıfa Göre")
                {

                    tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where sinifi Like '%" + txtSort.Text + "%'", bag);
                    adtr.Fill(tablo);
                    dataGridView1.DataSource = tablo;

                }

                else if (comboBox1.Text == "No'ya Göre")
                {

                    tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where ogrNo Like '%" + txtSort.Text + "%'", bag);
                    adtr.Fill(tablo);
                    dataGridView1.DataSource = tablo;

                }

                else if (comboBox1.Text == "Seri No'ya Göre")
                {

                    tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where tabletSeriNo Like '%" + txtSort.Text + "%'", bag);
                    adtr.Fill(tablo);
                    dataGridView1.DataSource = tablo;

                }


                else if (comboBox1.Text == "Soruna Göre")
                {

                    tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where tabletSorun Like '%" + txtSort.Text + "%'", bag);
                    adtr.Fill(tablo);
                    dataGridView1.DataSource = tablo;

                }

                else if (comboBox1.Text == "Tarihe Göre")
                {

                    tablo.Clear();
                    OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From ogrbir where tarih Like '%" + txtSort.Text + "%'", bag);
                    adtr.Fill(tablo);
                    dataGridView1.DataSource = tablo;

                }

            }
        }

        private void dosyaToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void exceleÇıkartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel.Application excelAktar = new Excel.Application();
            excelAktar.Visible = true;
            Excel.Workbook workbook = excelAktar.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet sheet1 = (Excel.Worksheet)excelAktar.ActiveSheet;
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Excel.Range myRange = (Excel.Range)sheet1.Cells[1, j + 1];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                myRange.ColumnWidth = 18;
                //   myRange.Columns.Font.Color = ColorTranslator.ToOle(Color.Red) ;
                myRange.Columns.Font.Bold = true;
                myRange.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            StartRow++;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    try
                    {
                        Excel.Range myRange = (Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];

                        string ss = dataGridView1[j, i].Value.ToString();
                        object sss = dataGridView1[j, i].Value;
                        myRange.Value2 = dataGridView1[j, i].Value == null || dataGridView1[j, i].Value is bool ? (j == 5 ? (bool)dataGridView1[j, i].Value == false ? "Hayır" : "Evet" : (bool)dataGridView1[j, i].Value == true ? "Arızalı" : "Sağlam") : (dataGridView1[j, i].Value);
                        myRange.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        private void yenileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listele();
        }

        private void hakkındaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Aziz Utku Kağıtcı","Hakkında",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private void yeniÖğrenciEkleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmEk = new frmEkle();
            frmEk.ShowDialog();
        }

        private void öğrenciDüzenleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmDuzen = new frmDuzenle();
            frmDuzenle.sAdi = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            frmDuzenle.sSinif = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            frmDuzenle.sNo = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            frmDuzenle.sSeriNo = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            frmDuzenle.cAriza = (bool)dataGridView1.CurrentRow.Cells[4].Value;
            frmDuzenle.cGaranti = (bool)dataGridView1.CurrentRow.Cells[5].Value;
            frmDuzenle.sSorun = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            frmDuzenle.sTarih = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            frmDuzen.ShowDialog();
        }

        private void öğrenciSilToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult cevap;
                cevap = MessageBox.Show("Kaydı silmek istediğinizden emin misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes)
                {
                    bag.Open();
                    kmt.Connection = bag;
                    kmt.CommandText = "DELETE FROM ogrbir WHERE ogrNo='" + dataGridView1.CurrentRow.Cells[2].Value.ToString() + "'";
                    kmt.ExecuteNonQuery();
                    bag.Close();
                    listele();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                bag.Close();
            }
        }

        private void çıkışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            String name = "Items";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "C:\\Sample.xlsx" +
                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
            con.Open();
            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);
            dataGridView1.DataSource = data;
        }
    }
}
