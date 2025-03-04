using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using UygunOlmayan.MyDb;
using UygunOlmayan.Tables;

namespace UygunOlmayan
{
    public partial class SorunBildir : Form
    {
        private MyDbContext dbContext;
        public SorunBildir()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string programadi = textBox1.Text;
            string baslık = textBox2.Text;
            string acıklama = textBox3.Text;

                var yeniSorun = new Sorun_Bildirim
                {
                    Sorun_Program = textBox1.Text,
                    Sorun_Baslık = textBox2.Text,
                    Sorun_Nedir = textBox3.Text
                };

                dbContext.sorun_Bildirims.Add(yeniSorun);
                dbContext.SaveChanges();

            MessageBox.Show("Sorun başarıyla bildirildi.");


        }
    }
}
