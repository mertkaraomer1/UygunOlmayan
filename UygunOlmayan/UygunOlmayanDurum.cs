using System.Data;
using System.Diagnostics;
using UygunOlmayan.MyDb;
using System.Net.Mail;
using System.Net;
using OfficeOpenXml;
using System.Drawing.Imaging;
using System.Drawing.Printing;


namespace UygunOlmayan.Tables
{
    public partial class UygunOlmayanDurum : Form
    {
        private static readonly HttpClient client = new HttpClient();
        private MyDbContext dbContext;
        private MySDbContext SdbContext;
        private PrintDocument printDocument;
        public UygunOlmayanDurum()
        {
            dbContext = new MyDbContext();
            SdbContext = new MySDbContext();
            InitializeComponent();
            printDocument = new PrintDocument();
            printDocument.PrintPage += new PrintPageEventHandler(PrintDocument_PrintPage);
        }
        public byte[] secilenResimBytes; // Resmi saklamak için genel değişken
        public string UrunID1
        {
            get { return textBox16.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox16.Text = value; }
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
        public string toplamMik1
        {
            get { return textBox17.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox17.Text = value; }
        }
        public string KayıpZaman1
        {
            get { return textBox7.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox7.Text = value; }
        }
        public string Hatatipi1
        {
            get { return comboBox2.Text; }
            set { comboBox2.Text = value; }
        }
        public string Acıklama1
        {
            get { return textBox9.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox9.Text = value; }
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

        public string Degerlendiren
        {
            get { return textBox6.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox6.Text = value; }
        }
        public string KokNedenAksiyon
        {
            get { return textBox14.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { textBox14.Text = value; }
        }

        public DateTime TerminTarihi
        {
            get { return dateTimePicker1.Value; } // DateTime olarak döndür
            set { dateTimePicker1.Value = value; } // DateTime olarak ayarla
        }

        public Image PictureBoxImage
        {
            get { return pictureBox1.Image; } // Image olarak döndür
            set
            {
                if (value != null)
                {
                    pictureBox1.Image = value; // Image olarak ayarla
                }
            }
        }
        public string uruntipi1
        {
            get { return comboBox1.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { comboBox1.Text = value; }
        }
        public string DuzelticiFaliyetDurum1
        {
            get { return comboBox5.Text; } // textBox1 burada TextBox'ın adı olmalı
            set { comboBox5.Text = value; }
        }
        public void LoadImageFromBytes(byte[] imageData)
        {
            if (imageData != null && imageData.Length > 0)
            {
                using (var ms = new MemoryStream(imageData))
                {
                    PictureBoxImage = Image.FromStream(ms); // Byte dizisini resme çevir ve PictureBox'a ata
                }
            }
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
                    toplamMiktar = Convert.ToInt32(textBox17.Text),
                    Tarih = DateTime.Now, 
                    KayıpZaman = Convert.ToInt32(textBox7.Text),
                    ZamanCinsi = "DK",
                    HataTipi = comboBox2.Text,
                    Aciklama = textBox9.Text,
                    Ozet = textBox8.Text,
                    HataBolumu = comboBox3.Text,
                    RaporuHazirlayan = textBox11.Text,
                    HatayıBulanBirim = comboBox4.Text,
                    KokNeden = textBox5.Text,
                    Aksiyon = textBox10.Text,
                    Sonuc = textBox12.Text,
                    Durum = "True",
                    Tedarikci = textBox13.Text,
                    Degerlendiren = textBox6.Text,
                    KokNedenAksiyon = textBox14.Text,
                    Resim = secilenResimBytes,
                    KapanısTarihi = new DateTime(1980, 5, 1),
                    TerminTarihi = dateTimePicker1.Value.Date,
                    urunimza = new Guid(),
                    uruntipi = comboBox1.Text,
                    DuzelticiFaliyetDurum = comboBox5.Text,

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
                textBox7.Clear();
                textBox8.Clear();
                textBox9.Clear();
                textBox10.Clear();
                textBox11.Clear();
                textBox12.Clear();
                textBox13.Clear();
                textBox14.Clear();
                textBox17.Clear();


            }
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

        }


        private void button4_Click(object sender, EventArgs e)
        {
            // UrunKodu ve UrunAdi'ye göre eşleşen HataliUrun nesnesini bulun
            var eskiHataliUrun = dbContext.hataliUruns
                .FirstOrDefault(u => u.UrunKodu == textBox1.Text && u.SiparisNo == textBox3.Text);

            if (eskiHataliUrun != null)
            {
                eskiHataliUrun.UrunKodu = textBox1.Text;
                eskiHataliUrun.UrunAdi = textBox2.Text;
                eskiHataliUrun.SiparisNo = textBox3.Text;
                eskiHataliUrun.HatalıMiktar = Convert.ToInt32(textBox4.Text);
                eskiHataliUrun.Adet = "Adet";
                eskiHataliUrun.toplamMiktar = Convert.ToInt32(textBox17.Text);
                eskiHataliUrun.KayıpZaman = Convert.ToInt32(textBox7.Text);
                eskiHataliUrun.ZamanCinsi = "DK";
                eskiHataliUrun.HataTipi = comboBox2.Text;
                eskiHataliUrun.Aciklama = textBox9.Text;
                eskiHataliUrun.Ozet = textBox8.Text;
                eskiHataliUrun.HataBolumu = comboBox3.Text;
                eskiHataliUrun.RaporuHazirlayan = textBox11.Text;
                eskiHataliUrun.HatayıBulanBirim = comboBox4.Text;
                eskiHataliUrun.KokNeden = textBox5.Text;
                eskiHataliUrun.Aksiyon = textBox10.Text;
                eskiHataliUrun.Sonuc = textBox12.Text;
                eskiHataliUrun.Durum = "False";
                eskiHataliUrun.Tedarikci = textBox13.Text;
                eskiHataliUrun.Degerlendiren = textBox6.Text;
                eskiHataliUrun.KokNedenAksiyon = textBox14.Text;
                eskiHataliUrun.TerminTarihi = dateTimePicker1.Value.Date;
                eskiHataliUrun.urunimza = new Guid();
                eskiHataliUrun.uruntipi = comboBox1.Text;
                eskiHataliUrun.DuzelticiFaliyetDurum = comboBox5.Text;

                // Eğer pictureBox1'de resim varsa, onu byte[] formatına çevir ve ata
                if (pictureBox1.Image != null)
                {
                    eskiHataliUrun.Resim = PictureBoxToByteArray(pictureBox1);
                }

                // Veritabanına değişiklikleri kaydet
                dbContext.SaveChanges();
                MessageBox.Show("Güncellendi.");
            }
        }

        private byte[] PictureBoxToByteArray(PictureBox pictureBox)
        {
            if (pictureBox.Image == null)
                return null;

            using (MemoryStream ms = new MemoryStream())
            {
                // Yeni bir bitmap oluşturarak resmi yeniden çiz (GDI+ hatasını önlemek için)
                using (Bitmap bmp = new Bitmap(pictureBox.Image))
                {
                    bmp.Save(ms, ImageFormat.Png);  // PNG formatı kullan
                }

                return ms.ToArray();
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


        private void eXCELÇEKToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExceliDoldurVeAc();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Resim Dosyaları|*.jpg;*.jpeg;*.png;*.bmp";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Seçilen resmi PictureBox'a yükle
                    Image selectedImage = Image.FromFile(openFileDialog.FileName);
                    pictureBox1.Image = selectedImage;

                    // Resmi byte[] dizisine çevir ve değişkene ata
                    secilenResimBytes = ResmiByteArrayeDonustur(selectedImage);

                }
            }
        }
        private byte[] ResmiByteArrayeDonustur(Image image)
        {
            if (image == null) return null;

            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg); // PNG veya BMP olabilir
                return ms.ToArray();
            }
        }

        private void ePOSTAGÖNDERToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string subject = "UYGUN OLMAYAN ÜRÜN KONTROL FORMU";
            string urunId = textBox16.Text;
            string body;

            // Bölümlere göre birden fazla e-posta adresi içeren sözlük tanımlandı.
            var emailAddresses = new Dictionary<string, List<string>>
                {
                    { "Montaj", new List<string> { "dturkan@icmmakina.com", "oocak@icmmakina.com" } },
                    { "Tasarım", new List<string> { "mbayram@icmmakina.com", "uulusoy@icmmakina.com" } },
                    { "İmalat", new List<string> { "pyesilyurt@icmmakina.com", "skoca@icmmakina.com", "hetanta@icmmakina.com" } },
                    { "Otomasyon", new List<string> { "otomasyon.proje@icmmakina.com", "tozpinar@icmmakina.com", "egozluk@icmmakina.com" } },
                    { "Satınalma", new List<string> { "satinalma@icmmakina.com" } },
                    { "Planlama", new List<string> { "shaci@icmmakina.com", "sbuyukay@icmmakina.com" } },
                    { "Kalite Kontrol", new List<string> { "oarslan@icmmakina.com" } },
                    { "Satış Sonrası", new List<string> { "hsokmen@icmmakina.com", "dtacyildiz@icmmakina.com" } },
                    { "Muhasebe", new List<string> { "bozcan@icmmakina.com", "mcelik@icmmakina.com" } }, 
                    { "Fabrika Müdürü", new List<string> { "ddeniz@icmmakina.com" } }
                };

            // Seçilen bölüme göre e-posta adresleri belirleniyor.
            if (emailAddresses.TryGetValue(comboBox3.Text, out List<string> emailList))
            {
                body = $"{urunId} NO'LU ÜRÜNÜN UYGUN OLMAYAN FORMUNU KONTROL EDİNİZ.";

                try
                {
                    // E-posta adreslerini virgülle ayırarak tek bir string olarak oluşturuyoruz.
                    string emailTo = string.Join(",", emailList);

                    SendEmail(emailTo, subject, body);
                    MessageBox.Show("E-posta başarıyla gönderildi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"E-posta gönderilemedi: {ex.Message}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Lütfen geçerli bir hata bölümü seçin!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


        }
        private void SendEmail(string to, string subject, string body)
        {
            MailMessage mail = new MailMessage();
            SmtpClient smtpServer = new SmtpClient("smtp.office365.com");

            mail.From = new MailAddress("oarslan@icmmakina.com"); // Gönderen e-posta adresi
            mail.To.Add(to);
            mail.Subject = subject;
            mail.Body = body;


            smtpServer.Port = 587; // Genellikle 587 veya 465
            smtpServer.Credentials = new NetworkCredential("oarslan@icmmakina.com", "B/132269177480ar"); // Şifreyi buraya ekleyin
            smtpServer.EnableSsl = true;

            smtpServer.Send(mail);
        }
        private void ExceliDoldurVeAc()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                // Excel dosyasının tam yolunu belirle
                string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Dosyalar", "UygunOlmayanUrunRaporu.xlsx");

                // Dosyanın var olup olmadığını kontrol et
                if (!File.Exists(excelFilePath))
                {
                    MessageBox.Show("Excel dosyası bulunamadı!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Excel dosyasını aç
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    // Çalışma sayfasını al
                    var worksheet = package.Workbook.Worksheets["FR.87.01 (R1)"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Çalışma sayfası bulunamadı!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // TextBox'lardan gelen verileri metin formatına dönüştür
                    string urunKodu = textBox1.Text;
                    string urunAdi = textBox2.Text;
                    string projekodu = textBox3.Text;
                    string hatalıMiktar = textBox4.Text;
                    string toplammiktar = textBox17.Text;
                    string islemKayipZamani = textBox7.Text + " DK";
                    string tespitEdenBolum = comboBox4.Text;
                    string raporuhazırlayan = textBox11.Text;
                    string hatanıntanımıAçıklama = textBox9.Text;
                    string kökneden = textBox5.Text;
                    string Aksiyon = textBox10.Text;
                    string sonuc = textBox8.Text;
                    string hatanınoluştuğubölüm = comboBox3.Text;
                    string ürüntipi = comboBox1.Text;
                    string düzelticifaliyet = comboBox5.Text;

                    if (ürüntipi == "Ham Malzeme")
                    {
                        worksheet.Cells["A9"].Value = "X"; // A9 hücresine değeri yaz

                    }
                    else if (ürüntipi == "Standart Malzeme")
                    {
                        worksheet.Cells["D9"].Value = "X"; // A9 hücresine değeri yaz
                    }
                    else if (ürüntipi == "Yarı Mamul")
                    {
                        worksheet.Cells["H9"].Value = "X"; // A9 hücresine değeri yaz
                    }
                    else if (ürüntipi == "Bitmiş Ürün")
                    {
                        worksheet.Cells["K9"].Value = "X"; // A9 hücresine değeri yaz
                    }

                    // Excel dosyasına yaz
                    worksheet.Cells["A4"].Value = $"Ürün Kodu: {urunKodu}";
                    worksheet.Cells["E4"].Value = $"Ürün Adı: {urunAdi}";
                    worksheet.Cells["J4"].Value = $"Sipariş No/Proje Kodu: {projekodu}";
                    worksheet.Cells["A5"].Value = $"Hatalı Miktar: {hatalıMiktar}";
                    worksheet.Cells["E5"].Value = $"Toplam Miktar: {toplammiktar}";
                    worksheet.Cells["K5"].Value = $"İşlem Kayıp Zamanı: {islemKayipZamani}";
                    worksheet.Cells["A6"].Value = $"Tespit Eden Bölüm: {tespitEdenBolum}";
                    worksheet.Cells["A7"].Value = $"Dış Tedarikçi veya Uygun Olmayan Ürünün Oluştuğu Bölüm :  {hatanınoluştuğubölüm}";
                    worksheet.Cells["A25"].Value = $"Raporu Hazırlayan: {raporuhazırlayan}";
                    // Aksiyon metnini hücreye ata
                    worksheet.Cells["A16"].Value = $"Hatanın Tanımı: {hatanıntanımıAçıklama}";
                    // Aksiyon metnini hücreye ata
                    worksheet.Cells["A21"].Value = $"Hatanın Kök Nedeni: {kökneden}";
                    // Aksiyon metnini hücreye ata
                    worksheet.Cells["A29"].Value = $"DEĞERLENDİRME - YAPILACAK FAALİYETLER: {Aksiyon}";
                    // Aksiyon metnini hücreye ata
                    worksheet.Cells["A41"].Value = $"DEĞERLENDİRME NOTLARI / SONUÇ: {sonuc}";
                    worksheet.Cells["A46"].Value = $"Düzeltici Faaliyet Gerekiyor: {düzelticifaliyet}";
                    // Dosyayı kaydet
                    package.Save();
                }

                // Dosyayı aç
                Process.Start(new ProcessStartInfo(excelFilePath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Bir hata oluştu: {ex.Message}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void yAZDIRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (PrintDialog printDialog = new PrintDialog())
            {
                printDialog.Document = printDocument;

                // Yazıcıyı seçme
                if (printDialog.ShowDialog() == DialogResult.OK)
                {
                    // Seçilen yazıcı ile yazdırma işlemini başlat
                    printDocument.Print();
                }
            }
        }

        private void lİSTEToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            UygunOlmayanListe ul = new UygunOlmayanListe();
            ul.Show();
            this.Hide();
        }

        private void rAPORToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UygunOlmayanRapor ur = new UygunOlmayanRapor();
            ur.Show();
        }

        private bool isZoomed = false;
        private Size originalSize;
        private Point originalLocation;
        private PictureBoxSizeMode originalSizeMode;
        private Control parentControl;


        private void pictureBox1_DoubleClick(object sender, EventArgs e)
        {
            if (!isZoomed)
            {
                // Mevcut ayarları kaydet
                originalSize = pictureBox1.Size;
                originalLocation = pictureBox1.Location;
                originalSizeMode = pictureBox1.SizeMode;
                parentControl = pictureBox1.Parent;

                // Form'un içine al ve en üste getir
                this.Controls.Add(pictureBox1);
                pictureBox1.BringToFront();

                // Resmi tam ekran yap
                pictureBox1.Dock = DockStyle.Fill;
                pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            }
            else
            {
                // Eski konumuna geri al
                pictureBox1.Dock = DockStyle.None;
                pictureBox1.Size = originalSize;
                pictureBox1.Location = originalLocation;
                pictureBox1.SizeMode = originalSizeMode;

                // Önceki ebeveyn kontrolüne geri taşı
                parentControl.Controls.Add(pictureBox1);
            }

            isZoomed = !isZoomed;

        }
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                e.Graphics.DrawImage(pictureBox1.Image, 0, 0, e.PageBounds.Width, e.PageBounds.Height);
            }
        }
    }
}
