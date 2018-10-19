using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Tablet_Takip_Programı
{
    public partial class frmGiris : Form
    {
        public Form1 frm1;
        public frmGiris()
        {
            InitializeComponent();
            frm1 = new Form1();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox1.Text.ToLower()=="azizutku" && textBox2.Text=="admin")
            {
                Form1.kullaniciadi = textBox1.Text.ToLower();
                frm1.Show();
                this.Hide();
                MessageBox.Show("Aziz Utku Kağıtcı Hoşgeldiniz!");
            }

            else if (textBox1.Text.ToLower() == "nesehayta" && textBox2.Text == "admin")
            {
                Form1.kullaniciadi = textBox1.Text.ToLower();
                frm1.Show();
                this.Hide();
                MessageBox.Show("Neşe Hayta Hoşgeldiniz!");
            }

            else if (textBox1.Text.ToLower() == "misafir" && textBox2.Text == "misafir")
            {
                Form1.kullaniciadi = textBox1.Text.ToLower();
                frm1.Show();
                this.Hide();
                MessageBox.Show("Misafir Hoşgeldiniz!");
            }

            else
            {
                MessageBox.Show("Kullanıcı Adı veya Şifre Yanlış!");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        double opa = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            opa = this.Opacity;
            opa = opa + 0.02;
            this.Opacity = opa;
        }

        private void frmGiris_Load(object sender, EventArgs e)
        {
            MessageBox.Show("Kullanıcı adı: misafir\nŞifre: misafir");
            timer1.Enabled = true;
        }
    }
}
