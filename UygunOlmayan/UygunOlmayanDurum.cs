using System.Data;
using UygunOlmayan.MyDb;

namespace UygunOlmayan.Tables
{
    public partial class UygunOlmayanDurum : Form
    {
        private MyDbContext dbContext;
        private MySDbContext SdbContext;
        public UygunOlmayanDurum()
        {
            dbContext = new MyDbContext();
            SdbContext=new MySDbContext();
            InitializeComponent();
        }
        public string UrunKodu1
        {
            get { return textBox1.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox1.Text = value; }
        }
        public string UrunAdi1
        {
            get { return textBox2.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox2.Text = value; }
        }
        public string SipNo1
        {
            get { return textBox3.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox3.Text = value; }
        }
        public string HataMik1
        {
            get { return textBox4.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox4.Text = value; }
        }
        public string KayıpZaman1
        {
            get { return textBox7.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox7.Text = value; }
        }
        public string Hatatipi1
        {
            get { return comboBox2.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { comboBox2.Text = value; }
        }
        public string Acıklama1
        {
            get { return textBox9.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox9.Text = value; }
        }
        public string Resim1
        {
            get { return textBox6.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox6.Text = value; }
        }
        public string ozet1
        {
            get { return textBox8.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox8.Text = value; }
        }
        public string HataBolumu1
        {
            get { return comboBox3.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { comboBox3.Text = value; }
        }
        public string Rhazirlayan1
        {
            get { return textBox11.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox11.Text = value; }
        }

        public string HataBulanBlum1
        {
            get { return comboBox4.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { comboBox4.Text = value; }
        }

        public string KokNeden1
        {
            get { return textBox5.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox5.Text = value; }
        }
        public string Aksiyon1
        {
            get { return textBox10.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox10.Text = value; }
        }
        public string Sonuc1
        {
            get { return textBox12.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox12.Text = value; }
        }
        public string tedarikci
        {
            get { return textBox13.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox13.Text = value; }
        }
        public bool ButtonGuncelle
        {
            get { return button4.Visible; }
            set { button4.Visible = value; }
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
                    Adet = "Adet",
                    Tarih = DateTime.Now, // Tarih alanı için uygun bir değer atayın
                    KayıpZaman = Convert.ToInt32(textBox7.Text),
                    ZamanCinsi = "DK",
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
                    Durum = "True",
                    Tedarikci=textBox13.Text,
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
                textBox13.Clear();

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

            if (selectedHataTipi == "Direk Satınalma-Component" || selectedHataTipi == "Fason Satınalma")
            {
                groupBox17.Visible = true;
            }
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
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // UrunKodu ve UrunAdi'ye göre eşleşen HataliUrun nesnesini bulun
            var eskiHataliUrun = dbContext.hataliUruns
                .FirstOrDefault(u => u.UrunKodu == textBox1.Text && u.SiparisNo == textBox3.Text);

            // Eğer eşleşen bir nesne bulunduysa, alanları güncelleyin
            if (eskiHataliUrun != null)
            {
                eskiHataliUrun.UrunKodu = textBox1.Text;
                eskiHataliUrun.UrunAdi = textBox2.Text;
                eskiHataliUrun.SiparisNo = textBox3.Text;
                eskiHataliUrun.HatalıMiktar = Convert.ToInt32(textBox4.Text);
                eskiHataliUrun.Adet = "Adet";
                eskiHataliUrun.KayıpZaman = Convert.ToInt32(textBox7.Text);
                eskiHataliUrun.ZamanCinsi = "DK";
                eskiHataliUrun.HataTipi = comboBox2.Text;
                eskiHataliUrun.Aciklama = textBox9.Text;
                eskiHataliUrun.Ozet = textBox8.Text;
                eskiHataliUrun.HataBolumu = comboBox3.Text;
                eskiHataliUrun.RaporuHazirlayan = textBox11.Text;
                eskiHataliUrun.HatayıBulanBirim = comboBox4.Text;
                eskiHataliUrun.Resim = textBox6.Text;
                eskiHataliUrun.KokNeden = textBox5.Text;
                eskiHataliUrun.Aksiyon = textBox10.Text;
                eskiHataliUrun.Sonuc = textBox12.Text;
                eskiHataliUrun.Durum = "True";

                // Değişiklikleri veritabanına kaydedin
                dbContext.SaveChanges();
                MessageBox.Show("Güncellendi.");
            }
        }

        private void UygunOlmayanDurum_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Form kapatılmak istendiğinde
            if (e.CloseReason == CloseReason.UserClosing)
            {
                // Formun kapatılmasını engelle
                e.Cancel = true;

                // Programı kapat
                Application.Exit();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string kod = textBox1.Text;

            // Veritabanından kod ile eşleşen ürünü buluyoruz
            var product = SdbContext.STOKLAR
                .FirstOrDefault(p => p.sto_kod == kod);

            if (product != null)
            {
                textBox2.Text = product.sto_isim;
            }
            else
            {
                textBox2.Text = "Kod bulunamadı";
            }
        }
    }
}
