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
        private System.Drawing.Printing.PrintDocument printDocument;
        private System.Threading.Timer _lookupTimer; // Stok arama için timer
        public UygunOlmayanDurum()
        {
            dbContext = new MyDbContext();
            SdbContext = new MySDbContext();
            InitializeComponent();
            printDocument = new PrintDocument();
            printDocument.PrintPage += PrintDocument_PrintPage;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
                dbContext?.Dispose();
                SdbContext?.Dispose();
                printDocument?.Dispose();
                _lookupTimer?.Dispose();
            }
            base.Dispose(disposing);
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
        public bool ButtonKaydet
        {
            get { return button1.Visible; }
            set { button1.Visible = value; }
        }
        private async void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox3.Text))
            {
                MessageBox.Show("Ürün Kodu ve Sipariş No alanları boş bırakılamaz!", "Zorunlu Alan", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                if (string.IsNullOrWhiteSpace(textBox1.Text)) textBox1.Focus(); else textBox3.Focus();
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            try
            {
                using (var db = new MyDbContext())
                {
                    var yeniHataliUrun = new HataliUrun
                    {
                        UrunKodu = textBox1.Text.Trim(),
                        UrunAdi = textBox2.Text.Trim(),
                        SiparisNo = textBox3.Text.Trim(),
                        HatalıMiktar = int.TryParse(textBox4.Text, out int hataliMiktar) ? hataliMiktar : 0,
                        Adet = "Adet",
                        toplamMiktar = int.TryParse(textBox17.Text, out int toplamMiktar) ? toplamMiktar : 0,
                        Tarih = DateTime.Now,
                        KayıpZaman = int.TryParse(textBox7.Text, out int kayipZaman) ? kayipZaman : 0,
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
                        urunimza = Guid.NewGuid(),
                        uruntipi = comboBox1.Text,
                        DuzelticiFaliyetDurum = comboBox5.Text,
                        AksiyonAlındı = "False"
                    };

                    db.hataliUruns.Add(yeniHataliUrun);
                    await db.SaveChangesAsync();

                    textBox16.Text = yeniHataliUrun.UrunId.ToString();
                    MessageBox.Show("Başarıyla kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    TemizleGirisAlanlari();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Kayıt sırasında hata oluştu: {ex.Message}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void TemizleGirisAlanlari()
        {
            Action<Control.ControlCollection> func = null;
            func = (controls) =>
            {
                foreach (Control control in controls)
                    if (control is TextBoxBase)
                        control.Text = string.Empty;
                    else
                        func(control.Controls);
            };
            func(this.Controls);
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


        private async void button4_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            try
            {
                var idStr = textBox16.Text;
                if (!int.TryParse(idStr, out int id)) return;

                using (var db = new MyDbContext())
                {
                    var eskiHataliUrun = await db.hataliUruns.FindAsync(id);
                    if (eskiHataliUrun != null)
                    {
                        eskiHataliUrun.UrunKodu = textBox1.Text;
                        eskiHataliUrun.UrunAdi = textBox2.Text;
                        eskiHataliUrun.SiparisNo = textBox3.Text;
                        eskiHataliUrun.HatalıMiktar = int.TryParse(textBox4.Text, out int hm) ? hm : 0;
                        eskiHataliUrun.toplamMiktar = int.TryParse(textBox17.Text, out int tm) ? tm : 0;
                        eskiHataliUrun.KayıpZaman = int.TryParse(textBox7.Text, out int kz) ? kz : 0;
                        eskiHataliUrun.HataTipi = comboBox2.Text;
                        eskiHataliUrun.Aciklama = textBox9.Text;
                        eskiHataliUrun.Ozet = textBox8.Text;
                        eskiHataliUrun.HataBolumu = comboBox3.Text;
                        eskiHataliUrun.RaporuHazirlayan = textBox11.Text;
                        eskiHataliUrun.HatayıBulanBirim = comboBox4.Text;
                        eskiHataliUrun.KokNeden = textBox5.Text;
                        eskiHataliUrun.Aksiyon = textBox10.Text;
                        eskiHataliUrun.Sonuc = textBox12.Text;
                        eskiHataliUrun.Tedarikci = textBox13.Text;
                        eskiHataliUrun.Degerlendiren = textBox6.Text;
                        eskiHataliUrun.KokNedenAksiyon = textBox14.Text;
                        eskiHataliUrun.TerminTarihi = dateTimePicker1.Value.Date;
                        eskiHataliUrun.uruntipi = comboBox1.Text;
                        eskiHataliUrun.DuzelticiFaliyetDurum = comboBox5.Text;

                        if (pictureBox1.Image != null)
                            eskiHataliUrun.Resim = PictureBoxToByteArray(pictureBox1);

                        await db.SaveChangesAsync();
                        MessageBox.Show("Başarıyla güncellendi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Güncelleme hatası: {ex.Message}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
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

        private System.Threading.Timer _lookupTimer; // Declare the timer

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // Debouncing: Her tuş basımında değil, yazma durunca ara
            _lookupTimer?.Dispose();
            _lookupTimer = new System.Threading.Timer(async _ => await RunStokLookup(), null, 500, Timeout.Infinite);
        }

        private async Task RunStokLookup()
        {
            if (!this.IsHandleCreated) return;

            string kod = "";
            this.Invoke(new Action(() => kod = textBox1.Text.Trim()));

            if (string.IsNullOrEmpty(kod)) return;

            try
            {
                using (var db = new MySDbContext()) // Assuming MySDbContext is your context for STOKLAR
                {
                    var product = await db.STOKLAR
                        .AsNoTracking()
                        .FirstOrDefaultAsync(p => p.sto_kod == kod);

                    if (this.IsHandleCreated)
                    {
                        this.Invoke(new Action(() =>
                        {
                            textBox2.Text = product?.sto_isim ?? "Kod bulunamadı";
                        }));
                    }
                }
            }
            catch { /* Silently fail for lookup */ }
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

        private async void ePOSTAGÖNDERToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox16.Text))
            {
                MessageBox.Show("Lütfen kayıt yaptıktan sonra tekrar deneyiniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var emailAddresses = GetEmailAddresses();
            if (emailAddresses.TryGetValue(comboBox3.Text, out List<string> emailList))
            {
                string subject = "UYGUN OLMAYAN ÜRÜN KONTROL FORMU";
                string body = $"{textBox16.Text} NO'LU ÜRÜNÜN UYGUN OLMAYAN FORMUNU KONTROL EDİNİZ.";
                string emailTo = string.Join(",", emailList);

                try
                {
                    await SendEmailAsync(emailTo, subject, body);
                    MessageBox.Show("E-posta başarıyla gönderildi.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"E-posta gönderilemedi: {ex.Message}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Lütfen geçerli bir hata bölümü seçiniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private async Task SendEmailAsync(string to, string subject, string body)
        {
            using (var mail = new MailMessage())
            using (var smtp = new SmtpClient("smtp.office365.com"))
            {
                mail.From = new MailAddress("oarslan@icmmakina.com", "ICM Makina Kalite");
                foreach (var address in to.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    mail.To.Add(address.Trim());

                mail.Subject = subject;
                mail.Body = body;

                smtp.Port = 587;
                smtp.Credentials = new NetworkCredential("oarslan@icmmakina.com", "D&351975783333ad");
                smtp.EnableSsl = true;

                await smtp.SendMailAsync(mail);
            }
        }

        private Dictionary<string, List<string>> GetEmailAddresses()
        {
            return new Dictionary<string, List<string>>
            {
                { "Montaj", new List<string> { "dturkan@icmmakina.com", "oocak@icmmakina.com" } },
                { "Tasarım", new List<string> { "mbayram@icmmakina.com", "uulusoy@icmmakina.com" } },
                { "İmalat", new List<string> { "pyesilyurt@icmmakina.com", "skoca@icmmakina.com", "hetanta@icmmakina.com" } },
                { "Otomasyon", new List<string> { "otomasyon.proje@icmmakina.com", "tozpinar@icmmakina.com", "egozluk@icmmakina.com", "bguden@icmmakina.com" , "byanik@icmmakina.com" } },
                { "Satınalma", new List<string> { "satinalma@icmmakina.com" } },
                { "Planlama", new List<string> { "shaci@icmmakina.com", "sbuyukay@icmmakina.com" } },
                { "Kalite Kontrol", new List<string> { "oarslan@icmmakina.com", "bgebedek@icmmakina.com" } },
                { "Satış Sonrası", new List<string> { "hsokmen@icmmakina.com", "dtacyildiz@icmmakina.com" } },
                { "Muhasebe", new List<string> { "bozcan@icmmakina.com", "mcelik@icmmakina.com" } },
                { "Fabrika Müdürü", new List<string> { "ddeniz@icmmakina.com" } }
            };
        }
        private async void ExceliDoldurVeAc()
        {
            this.Cursor = Cursors.WaitCursor;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Dosyalar", "UygunOlmayanUrunRaporu.xlsx");

                if (!File.Exists(excelFilePath))
                {
                    MessageBox.Show("Excel şablon dosyası bulunamadı!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Geçici bir kopya üzerinde çalış (Dosya kullanımda hatası almamak için)
                string tempPath = Path.Combine(Path.GetTempPath(), $"Rapor_{Guid.NewGuid().ToString().Substring(0,8)}.xlsx");
                File.Copy(excelFilePath, tempPath, true);

                using (var package = new ExcelPackage(new FileInfo(tempPath)))
                {
                    var worksheet = package.Workbook.Worksheets["FR.87.01 (R1)"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("'FR.87.01 (R1)' isimli çalışma sayfası şablonda bulunamadı!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Hücreleri doldur
                    string ürüntipi = comboBox1.Text;
                    worksheet.Cells["A9,D9,H9,K9"].Value = ""; // Reset checkboxes
                    
                    if (ürüntipi == "Ham Malzeme") worksheet.Cells["A9"].Value = "X";
                    else if (ürüntipi == "Standart Malzeme") worksheet.Cells["D9"].Value = "X";
                    else if (ürüntipi == "Yarı Mamul") worksheet.Cells["H9"].Value = "X";
                    else if (ürüntipi == "Bitmiş Ürün") worksheet.Cells["K9"].Value = "X";

                    worksheet.Cells["A4"].Value = $"Ürün Kodu: {textBox1.Text}";
                    worksheet.Cells["E4"].Value = $"Ürün Adı: {textBox2.Text}";
                    worksheet.Cells["J4"].Value = $"Sipariş No/Proje Kodu: {textBox3.Text}";
                    worksheet.Cells["A5"].Value = $"Hatalı Miktar: {textBox4.Text}";
                    worksheet.Cells["E5"].Value = $"Toplam Miktar: {textBox17.Text}";
                    worksheet.Cells["K5"].Value = $"İşlem Kayıp Zamanı: {textBox7.Text} DK";
                    worksheet.Cells["A6"].Value = $"Tespit Eden Bölüm: {comboBox4.Text}";
                    worksheet.Cells["A7"].Value = $"Dış Tedarikçi veya Uygun Olmayan Ürünün Oluştuğu Bölüm : {comboBox3.Text}";
                    worksheet.Cells["A25"].Value = $"Raporu Hazırlayan: {textBox11.Text}";
                    worksheet.Cells["A16"].Value = $"Hatanın Tanımı: {textBox9.Text}";
                    worksheet.Cells["A21"].Value = $"Hatanın Kök Nedeni: {textBox5.Text}";
                    worksheet.Cells["A29"].Value = $"DEĞERLENDİRME - YAPILACAK FAALİYETLER: {textBox10.Text}";
                    worksheet.Cells["A41"].Value = $"DEĞERLENDİRME NOTLARI / SONUÇ: {textBox8.Text}";
                    worksheet.Cells["A46"].Value = $"Düzeltici Faaliyet Gerekiyor: {comboBox5.Text}";

                    await package.SaveAsync();
                }

                Process.Start(new ProcessStartInfo(tempPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excel raporu oluşturulurken bir hata oluştu: {ex.Message}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
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

        private void bİLDİRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SorunBildir SB = new SorunBildir();
            SB.Show();
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox16.Text))
            {
                MessageBox.Show("Önce kayıt yapmalısınız.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            try
            {
                var emailAddresses = GetEmailAddresses();
                if (emailAddresses.TryGetValue(comboBox6.Text, out List<string> emailList))
                {
                    string subject = "UYGUN OLMAYAN ÜRÜN KONTROL FORMU - AKSİYON TALEBİ";
                    string body = $"{textBox16.Text} NO'LU ÜRÜNÜN UYGUN OLMAYAN FORMUNU KONTROL EDİNİZ.";
                    string emailTo = string.Join(",", emailList);

                    await SendEmailAsync(emailTo, subject, body);

                    using (var db = new MyDbContext())
                    {
                        var bölüm = new AksiyonAlacakBölüm
                        {
                            UygunOlmayanId = int.Parse(textBox16.Text),
                            AksiyonBölümü = comboBox6.Text,
                            Tarihi = DateTime.Now
                        };

                        db.aksiyonAlacakBölüms.Add(bölüm);
                        await db.SaveChangesAsync();
                    }

                    MessageBox.Show("Aksiyon maili gönderildi ve kaydedildi.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Lütfen geçerli bir aksiyon bölümü seçiniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"İşlem sırasında hata oluştu: {ex.Message}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
        private void pROGRAMIKAPATToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Programı kapatmak istediğinize emin misiniz?",
                                                  "Çıkış Onayı",
                                                  MessageBoxButtons.YesNo,
                                                  MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void rAPORToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            RAPOR R=new RAPOR();
            R.Show();
        }
    }
}
