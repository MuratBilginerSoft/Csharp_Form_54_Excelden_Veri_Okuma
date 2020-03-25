using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb; // Eklenmesi Gereken Kütüphane

namespace Kod_H_Excelden_Veri_Okuma
{
    public partial class Form1 : Form
    {
        #region Değişkenler

        string DosyaYolu;
        string Sorgu;

        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            textBox1.Text = openFileDialog1.FileName;
            DosyaYolu = openFileDialog1.FileName;

            OleDbConnection Baglanti = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" + DosyaYolu + "; Extended Properties=Excel 12.0");

            if (Baglanti.State==ConnectionState.Closed)
            {
                Baglanti.Open();
            }

            else
            {
                Baglanti.Close();
                Baglanti.Open();
            }

            Sorgu = "Select * from [Sayfa1$]";
            OleDbDataAdapter Adapter = new OleDbDataAdapter(Sorgu, Baglanti);
            Baglanti.Close();

            DataTable VeriTablo = new DataTable();
            Adapter.Fill(VeriTablo);

            dataGridView1.DataSource = VeriTablo;
            
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
    }
}
