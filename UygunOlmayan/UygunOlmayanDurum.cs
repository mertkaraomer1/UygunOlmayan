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

namespace UygunOlmayan.Tables
{
    public partial class UygunOlmayanDurum : Form
    {
        private MyDbContext dbContext;
        public UygunOlmayanDurum()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 1. Veritabanı bağlamınızı oluşturun
            using (var dbContext = new MyDbContext())
            {
                // 2. HataliUrun sınıfı ile eşleşen bir nesne oluşturun ve verileri atayın
                var yeniHataliUrun = new HataliUrun
                {
                    UrunKodu = textBox1.Text,
                    UrunAdi = textBox2.Text,
                    SiparisNo = textBox3.Text,
                    HatalıMiktar = Convert.ToInt32(textBox4.Text),
                    Tarih = DateTime.Now, // Tarih alanı için uygun bir değer atayın
                    KayıpZaman = textBox7.Text,
                    HataTipi = comboBox2.Text,
                    Aciklama = textBox9.Text,
                    Ozet = textBox8.Text,
                    HataBolumu = comboBox3.Text,
                    RaporuHazirlayan = textBox11.Text,
                    HatayıBulanBirim = comboBox4.Text,
                    Resim = textBox6.Text,
                    KokNeden = textBox5.Text,
                    Aksiyon = textBox10.Text,
                    Sonuc = textBox12.Text,
                };

                // 3. Veritabanına ekleme işlemi
                dbContext.hataliUruns.Add(yeniHataliUrun);
                dbContext.SaveChanges(); // Değişiklikleri kaydedin
                MessageBox.Show("Kaydedildi...");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
                textBox6.Clear();
                textBox7.Clear();
                textBox8.Clear();
                textBox9.Clear();
                textBox10.Clear();
                textBox11.Clear();
                textBox12.Clear();


            }
        }
        private void UygunOlmayanDurum_Load(object sender, EventArgs e)
        {
            var hatatipi = dbContext.hataGruplars.Select(a => a.HataTipi).ToList();
            comboBox2.DataSource = hatatipi;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            pictureBox1.ImageLocation = openFileDialog1.FileName;
            textBox6.Text = openFileDialog1.FileName;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            // ComboBox'dan seçilen HataTipi'ni alın
            string selectedHataTipi = comboBox2.SelectedItem.ToString();

            // HataTipi'ne göre verileri seçin
            var hataliUrunler = dbContext.hataGruplars.Where(h => h.HataTipi == selectedHataTipi).ToList();


            // İlk veriyi seçildiğinde TextBox9'a HataAcıklaması'nı yazdırın (varsa)
            if (hataliUrunler.Count > 0)
            {
                string hataAciklamasi = hataliUrunler[0].HataAcıklaması; // İlk verinin HataAciklamasi
                textBox9.Text = hataAciklamasi;
            }
            else
            {
                textBox9.Text = ""; // Veri bulunamazsa TextBox'i boşalt
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UygunOlmayanListe ul = new UygunOlmayanListe();
            ul.Show();
        }
    }
}
