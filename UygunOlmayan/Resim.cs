using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UygunOlmayan
{
    public partial class Resim : Form
    {
        public Resim()
        {
            InitializeComponent();
        }
        public string picturebox132
        {
            get { return pictureBox1.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { pictureBox1.Text = value; }
        }
        private void Resim_Load(object sender, EventArgs e)
        {
            string resimYolu = picturebox132; // picturebox, dosyanın yolunu içeriyor

            if (!string.IsNullOrEmpty(resimYolu) && File.Exists(resimYolu))
            {
                pictureBox1.Image = Image.FromFile(resimYolu);
            }
            else
            {
                // Dosya yok veya yol hatalıysa bir varsayılan resim veya bir hata mesajı gösterebilirsiniz.
                // Örnek: pictureBox1.Image = Properties.Resources.VarsayilanResim;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Resim res = new Resim();
            this.Close();
        }

        private void Resim_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }
    }
}
