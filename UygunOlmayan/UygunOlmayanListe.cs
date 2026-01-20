using System.Data;
using System.Net.Mail;
using System.Net;
using System.Windows.Forms;
using UygunOlmayan.MyDb;
using UygunOlmayan.Tables;
using OfficeOpenXml;
using System.Diagnostics;
using Microsoft.EntityFrameworkCore;

namespace UygunOlmayan
{
    public partial class UygunOlmayanListe : Form
    {
        MyDbContext dbContext;
        private UygunOlmayanDurum UOD;
        public UygunOlmayanListe()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
            EnableDoubleBuffered(advancedDataGridView1);
        }

        private void EnableDoubleBuffered(Control control)
        {
            var property = typeof(Control).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            property?.SetValue(control, true, null);
        }


        DataTable table = new DataTable();
        public async Task UygunOlmayanVol1Async()
        {
            table.Clear();
            if (table.Columns.Count == 0)
            {
                string[] columnNames = {
                    "Urun Id", "Urun Kodu", "Urun Adi", "Siparis No", "Hatalı Miktar", "Adet",
                    "Toplam Miktar", "Tarih", "Kayıp Zaman", "Zaman Türü", "Hata Tipi", "Aciklama",
                    "Tedarikçi", "Ozet", "Hata Bolumu", "Raporu Hazirlayan", "Hatayı Bulan Birim",
                    "Kök Neden", "Aksiyon", "Sonuç", "Değerlendiren", "Kök Neden Aksiyon", "ResimVar",
                    "Kapanış Tarihi", "Termin Tarihi", "Ürün Tipi", "Düzeltici Faaliyet Var Mı?", "Durum", "Aksiyon Alancak Bölüm"
                };
                foreach (string columnName in columnNames) table.Columns.Add(columnName);
            }

            try
            {
                using (var db = new MyDbContext())
                {
                    // 🔥 HIZLANDIRMA: Projeksiyon (.Select) kullanarak devasa resim datalarını (BLOB) çekmiyoruz.
                    // Sadece ekranda görünen kolonları seçiyoruz.
                    var hataliUrunSorgu = db.hataliUruns
                        .Where(x => x.urunimza == Guid.Empty)
                        .OrderByDescending(x => x.Tarih)
                        .Select(u => new 
                        {
                            u.UrunId, u.UrunKodu, u.UrunAdi, u.SiparisNo, u.HatalıMiktar, u.Adet,
                            u.toplamMiktar, u.Tarih, u.KayıpZaman, u.ZamanCinsi,
                            u.HataTipi, u.Aciklama, u.Tedarikci, u.Ozet, u.HataBolumu, u.RaporuHazirlayan,
                            u.HatayıBulanBirim, u.KokNeden, u.Aksiyon, u.Sonuc, u.Degerlendiren,
                            u.KokNedenAksiyon, ResimVar = u.Resim != null ? "VAR" : "", 
                            u.KapanısTarihi, u.TerminTarihi, u.uruntipi, u.DuzelticiFaliyetDurum,
                            u.Durum, u.AksiyonAlındı
                        });

                    var hataliUrunList = await hataliUrunSorgu.ToListAsync();
                    var urunIds = hataliUrunList.Select(x => x.UrunId).ToList();

                    // Aksiyonları tek seferde çek
                    var aksiyonlar = db.aksiyonAlacakBölüms
                        .Where(a => urunIds.Contains(a.UygunOlmayanId))
                        .GroupBy(a => a.UygunOlmayanId)
                        .Select(g => g.OrderByDescending(x => x.Tarihi).FirstOrDefault())
                        .ToDictionary(a => a.UygunOlmayanId, a => a.AksiyonBölümü);

                    foreach (var urun in hataliUrunList)
                    {
                        string durumStr = (urun.Durum ?? "False") == "False" ? "İŞLEMİ DEVAM EDİYOR." : "DEĞERLENDİRMEYİ BEKLİYOR.";
                        string aksiyonAlındımı = (urun.AksiyonAlındı ?? "False") == "False" ? durumStr : "AKSİYON ALINDI";
                        aksiyonlar.TryGetValue(urun.UrunId, out string aksBolumu);

                        table.Rows.Add(
                            urun.UrunId, urun.UrunKodu, urun.UrunAdi, urun.SiparisNo, urun.HatalıMiktar, urun.Adet,
                            urun.toplamMiktar, urun.Tarih.ToString("yyyy.MM.dd"), urun.KayıpZaman, urun.ZamanCinsi,
                            urun.HataTipi, urun.Aciklama, urun.Tedarikci, urun.Ozet, urun.HataBolumu, urun.RaporuHazirlayan,
                            urun.HatayıBulanBirim, urun.KokNeden, urun.Aksiyon, urun.Sonuc, urun.Degerlendiren,
                            urun.KokNedenAksiyon, urun.ResimVar, urun.KapanısTarihi, urun.TerminTarihi,
                            urun.uruntipi, urun.DuzelticiFaliyetDurum, aksiyonAlındımı, aksBolumu ?? ""
                        );
                    }
                }

                advancedDataGridView1.DataSource = table;

                // 🔥 Performans İçin: Sürekli AutoSize yerine bir seferlik boyutlandırma yapıyoruz
                advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                advancedDataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                
                SetupGridStyles();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hızlı yükleme sırasında hata: {ex.Message}");
            }
        }

        private void SetupGridStyles()
        {
            // 1. Gereksiz sütunları kaldır ve buton sütunlarını isimle yönet
            if (advancedDataGridView1.Columns.Contains("ActionBtn")) advancedDataGridView1.Columns.Remove("ActionBtn");
            if (advancedDataGridView1.Columns.Contains("ImzaBtn")) advancedDataGridView1.Columns.Remove("ImzaBtn");

            // 2. Aksiyon Butonu (Modern İkonlu)
            DataGridViewImageColumn actionCol = new DataGridViewImageColumn
            {
                Name = "ActionBtn",
                HeaderText = "AKSİYON",
                Image = File.Exists("delete.png") ? Image.FromFile("delete.png") : null,
                ImageLayout = DataGridViewImageCellLayout.Zoom,
                Width = 70,
                ToolTipText = "Aksiyonu Tamamla"
            };
            advancedDataGridView1.Columns.Add(actionCol);

            // 3. İmza Butonu (Modern İkonlu)
            DataGridViewImageColumn imzaCol = new DataGridViewImageColumn
            {
                Name = "ImzaBtn",
                HeaderText = "İMZA",
                Image = File.Exists("imza.png") ? Image.FromFile("imza.png") : null,
                ImageLayout = DataGridViewImageCellLayout.Zoom,
                Width = 70,
                ToolTipText = "Yetkili İmza At"
            };
            advancedDataGridView1.Columns.Add(imzaCol);

            // 4. Genel Grid Ayarları (Kurumsal Görünüm)
            advancedDataGridView1.BorderStyle = BorderStyle.None;
            advancedDataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            advancedDataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            advancedDataGridView1.GridColor = Color.FromArgb(224, 224, 224);
            advancedDataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            advancedDataGridView1.BackgroundColor = Color.White;
            advancedDataGridView1.RowTemplate.Height = 40;

            // 5. Kurumsal Başlık Stili
            var headerStyle = new DataGridViewCellStyle
            {
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
                BackColor = Color.FromArgb(41, 57, 85),
                ForeColor = Color.White,
                SelectionBackColor = Color.FromArgb(41, 57, 85),
                Alignment = DataGridViewContentAlignment.MiddleCenter
            };
            advancedDataGridView1.EnableHeadersVisualStyles = false;
            advancedDataGridView1.ColumnHeadersDefaultCellStyle = headerStyle;
            advancedDataGridView1.ColumnHeadersHeight = 45;

            // 6. Hücre Stilleri ve Seçim Rengi
            var cellStyle = new DataGridViewCellStyle
            {
                Font = new Font("Segoe UI", 9F),
                SelectionBackColor = Color.FromArgb(226, 230, 236),
                SelectionForeColor = Color.Black
            };
            advancedDataGridView1.DefaultCellStyle = cellStyle;
            advancedDataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 251, 252);

            // 7. Durum Sütununu Kurumsal Renklerle Biçimlendir
            RenklendirDurum();
        }

        private void RenklendirDurum()
        {
            if (!advancedDataGridView1.Columns.Contains("Durum")) return;
            int colIndex = advancedDataGridView1.Columns["Durum"].Index;

            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.IsNewRow || row.Cells[colIndex].Value == null) continue;
                string val = row.Cells[colIndex].Value.ToString();

                if (val == "İŞLEMİ DEVAM EDİYOR.")
                {
                    row.Cells[colIndex].Style.BackColor = Color.FromArgb(255, 193, 7); // Warning (Amber)
                    row.Cells[colIndex].Style.ForeColor = Color.Black;
                }
                else if (val == "DEĞERLENDİRMEYİ BEKLİYOR.")
                {
                    row.Cells[colIndex].Style.BackColor = Color.FromArgb(220, 53, 69); // Danger (Red)
                    row.Cells[colIndex].Style.ForeColor = Color.White;
                }
                else if (val == "AKSİYON ALINDI")
                {
                    row.Cells[colIndex].Style.BackColor = Color.FromArgb(40, 167, 69); // Success (Green)
                    row.Cells[colIndex].Style.ForeColor = Color.White;
                }
            }
        }

        private async void UygunOlmayanListe_Load(object sender, EventArgs e)
        {
            await UygunOlmayanVol1Async();
        }

        private void advancedDataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // Hücre değerlerini isimle (by Name) alarak hatayı önlüyoruz
            var row = advancedDataGridView1.Rows[e.RowIndex];
            
            // Eğer UOD formu açıksa ve bertaraf edilmediyse (Disposed), mevcut formu getir veya yeniden oluştur
            if (UOD == null || UOD.IsDisposed)
            {
                UOD = new UygunOlmayanDurum();
            }
            UOD.BringToFront();

            try
            {
                string urunID = row.Cells["Urun Id"]?.Value?.ToString() ?? "0";
                
                // Formu doldurma - Null kontrolleriyle beraber
                UOD.UrunID1 = urunID;
                UOD.UrunKodu1 = row.Cells["Urun Kodu"]?.Value?.ToString() ?? "";
                UOD.UrunAdi1 = row.Cells["Urun Adi"]?.Value?.ToString() ?? "";
                UOD.SipNo1 = row.Cells["Siparis No"]?.Value?.ToString() ?? "";
                UOD.HataMik1 = row.Cells["Hatalı Miktar"]?.Value?.ToString() ?? "0";
                UOD.toplamMik1 = row.Cells["Toplam Miktar"]?.Value?.ToString() ?? "0";
                UOD.KayıpZaman1 = row.Cells["Kayıp Zaman"]?.Value?.ToString() ?? "0";
                UOD.Hatatipi1 = row.Cells["Hata Tipi"]?.Value?.ToString() ?? "";
                UOD.Acıklama1 = row.Cells["Aciklama"]?.Value?.ToString() ?? "";
                UOD.ozet1 = row.Cells["Ozet"]?.Value?.ToString() ?? "";
                UOD.HataBolumu1 = row.Cells["Hata Bolumu"]?.Value?.ToString() ?? "";
                UOD.Rhazirlayan1 = row.Cells["Raporu Hazirlayan"]?.Value?.ToString() ?? "";
                UOD.HataBulanBlum1 = row.Cells["Hatayı Bulan Birim"]?.Value?.ToString() ?? "";
                UOD.KokNeden1 = row.Cells["Kök Neden"]?.Value?.ToString() ?? "";
                UOD.Aksiyon1 = row.Cells["Aksiyon"]?.Value?.ToString() ?? "";
                UOD.Sonuc1 = row.Cells["Sonuç"]?.Value?.ToString() ?? "";
                UOD.tedarikci = row.Cells["Tedarikçi"]?.Value?.ToString() ?? "";
                UOD.Degerlendiren = row.Cells["Değerlendiren"]?.Value?.ToString() ?? "";
                UOD.KokNedenAksiyon = row.Cells["Kök Neden Aksiyon"]?.Value?.ToString() ?? "";
                UOD.uruntipi1 = row.Cells["Ürün Tipi"]?.Value?.ToString() ?? "";
                UOD.DuzelticiFaliyetDurum1 = row.Cells["Düzeltici Faaliyet Var Mı?"]?.Value?.ToString() ?? "";

                // Tarih dönüşümü
                if (DateTime.TryParse(row.Cells["Termin Tarihi"]?.Value?.ToString(), out DateTime terminTarihi))
                    UOD.TerminTarihi = terminTarihi;

                // Resmi sadece ihtiyaç duyulduğunda DB'den çek (Performans için)
                int id = int.Parse(urunID);
                using (var db = new MyDbContext())
                {
                    var resimData = db.hataliUruns.Where(x => x.UrunId == id).Select(x => x.Resim).FirstOrDefault();
                    UOD.LoadImageFromBytes(resimData);
                }

                UOD.ButtonGuncelle = true;
                UOD.ButtonKaydet = false;
                UOD.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Detaylar açılırken hata oluştu: {ex.Message}");
            }
        }

        private void advancedDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            string colName = advancedDataGridView1.Columns[e.ColumnIndex].Name;

            if (colName == "ImzaBtn") // Yetkili İmza Alanı (Eski index 30 civarı)
            {
                DataGridViewRow gridRow = advancedDataGridView1.Rows[e.RowIndex];

                int urunId;
                if (!int.TryParse(gridRow.Cells["Urun Id"]?.Value?.ToString(), out urunId))
                    return;

                var eskiHataliUrun = dbContext.hataliUruns
                    .FirstOrDefault(x => x.UrunId == urunId);

                if (eskiHataliUrun == null)
                    return;

                // 🔥 DB güncelle
                eskiHataliUrun.urunimza = Guid.NewGuid();
                eskiHataliUrun.KapanısTarihi = DateTime.Now.Date;
                dbContext.SaveChanges();

                // 🔥 DataTable'dan sil (Ekranda görünürse filtreden bağımsız kalkar)
                DataRowView drv = gridRow.DataBoundItem as DataRowView;
                if (drv != null)
                {
                    drv.Row.Delete();
                    table.AcceptChanges(); // Değişikliği onayla
                }

                MessageBox.Show("Ürün kapatıldı ve listeden kaldırıldı.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (colName == "ActionBtn") // Aksiyon Alındı Butonu
            {
                DataGridViewRow row = advancedDataGridView1.Rows[e.RowIndex];
                string urunKod = row.Cells["Urun Kodu"]?.Value?.ToString();
                string sipNo = row.Cells["Siparis No"]?.Value?.ToString();

                if (string.IsNullOrEmpty(urunKod) || string.IsNullOrEmpty(sipNo)) return;

                var eskiHataliUrun = dbContext.hataliUruns.FirstOrDefault(u => u.UrunKodu == urunKod && u.SiparisNo == sipNo);
                if (eskiHataliUrun == null) return;

                eskiHataliUrun.AksiyonAlındı = "True";
                dbContext.SaveChanges();

                // 🔥 Grid yerine doğrudan DataTable satırını güncelle
                // Böylece eğer bir filtre aktifse (örn: Durum != 'AKSİYON ALINDI'), satır anında ekrandan kalkar.
                DataRowView drv = row.DataBoundItem as DataRowView;
                if (drv != null)
                {
                    drv.Row["Durum"] = "AKSİYON ALINDI";
                    // Filtrenin yeniden tetiklenmesi için RowFilter'ı tazele
                    string currentFilter = table.DefaultView.RowFilter;
                    table.DefaultView.RowFilter = string.Empty;
                    table.DefaultView.RowFilter = currentFilter;
                }

                RenklendirDurum();
                MessageBox.Show("Aksiyon kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        
        private void lİSTEYİEXCELEAKTARToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportToExcelFast();
        }

        private void ExportToExcelFast()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("LİSTE");
                
                // Başlıkları aktar
                for (int i = 0; i < advancedDataGridView1.Columns.Count; i++)
                {
                    if (advancedDataGridView1.Columns[i] is DataGridViewImageColumn) continue;
                    worksheet.Cells[1, i + 1].Value = advancedDataGridView1.Columns[i].HeaderText;
                }

                // Verileri aktar
                for (int i = 0; i < advancedDataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < advancedDataGridView1.Columns.Count; j++)
                    {
                        if (advancedDataGridView1.Columns[j] is DataGridViewImageColumn) continue;
                        worksheet.Cells[i + 2, j + 1].Value = advancedDataGridView1.Rows[i].Cells[j].Value;
                    }
                }

                // Kaydet ve Aç
                string path = Path.Combine(Path.GetTempPath(), $"Liste_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                File.WriteAllBytes(path, package.GetAsByteArray());
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
        }

        private async void ePOSTAGÖNDERToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewRow selectedRow = advancedDataGridView1.CurrentRow;
            if (selectedRow == null)
            {
                MessageBox.Show("Lütfen bir satır seçin!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string urunId = selectedRow.Cells["Urun Id"].Value?.ToString();
            string hataBolumu = selectedRow.Cells["Hata Bolumu"].Value?.ToString();
            
            var emailAddresses = GetEmailDictionary();

            if (emailAddresses.TryGetValue(hataBolumu, out List<string> emailList))
            {
                string subject = "UYGUN OLMAYAN ÜRÜN KONTROL FORMU";
                string body = $"{urunId} NO'LU ÜRÜNÜN UYGUN OLMAYAN FORMUNU KONTROL EDİNİZ.";
                string emails = string.Join(",", emailList);

                try 
                {
                    await SendEmailAsync(emails, subject, body);
                    MessageBox.Show("E-posta başarıyla gönderildi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"E-posta hatası: {ex.Message}");
                }
            }
            else
            {
                MessageBox.Show("Bu bölüm için tanımlı e-posta bulunamadı.");
            }
        }

        private Dictionary<string, List<string>> GetEmailDictionary()
        {
            return new Dictionary<string, List<string>>
            {
                { "Montaj", new List<string> { "dturkan@icmmakina.com", "oocak@icmmakina.com" } },
                { "Tasarım", new List<string> { "mbayram@icmmakina.com", "uulusoy@icmmakina.com" } },
                { "İmalat", new List<string> { "pyesilyurt@icmmakina.com", "skoca@icmmakina.com", "hetanta@icmmakina.com" } },
                { "Otomasyon", new List<string> { "otomasyon.proje@icmmakina.com", "tozpinar@icmmakina.com", "egozluk@icmmakina.com", "bguden@icmmakina.com" , "byanik@icmmakina.com" } },
                { "Satınalma", new List<string> { "satinalma@icmmakina.com" } },
                { "Planlama", new List<string> { "shaci@icmmakina.com", "sbuyukay@icmmakina.com" } },
                { "Kalite Kontrol", new List<string> { "oarslan@icmmakina.com" } },
                { "Satış Sonrası", new List<string> { "hsokmen@icmmakina.com", "dtacyildiz@icmmakina.com" } },
                { "Muhasebe", new List<string> { "bozcan@icmmakina.com", "mcelik@icmmakina.com" } },
                { "Fabrika Müdürü", new List<string> { "ddeniz@icmmakina.com" } }
            };
        }

        private async Task SendEmailAsync(string to, string subject, string body)
        {
            using (MailMessage mail = new MailMessage())
            using (SmtpClient smtpServer = new SmtpClient("smtp.office365.com"))
            {
                mail.From = new MailAddress("oarslan@icmmakina.com", "ICM Makina Kalite");
                foreach (var address in to.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    mail.To.Add(address.Trim());

                mail.Subject = subject;
                mail.Body = body;

                smtpServer.Port = 587;
                smtpServer.Credentials = new NetworkCredential("oarslan@icmmakina.com", "D&351975783333ad");
                smtpServer.EnableSsl = true;

                await smtpServer.SendMailAsync(mail);
            }
        }

        private void eXCELÇEKToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewRow selectedRow = advancedDataGridView1.CurrentRow;
            if (selectedRow == null)
            {
                MessageBox.Show("Lütfen bir satır seçiniz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Seçili satırdan gerekli verileri al
            var urunIdValue = selectedRow.Cells["Urun Id"].Value?.ToString();
            var urunKodu = selectedRow.Cells["Urun Kodu"].Value?.ToString();
            var urunAdi = selectedRow.Cells["Urun Adi"].Value?.ToString();
            var siparisNo = selectedRow.Cells["Siparis No"].Value?.ToString();
            var hataliMiktar = selectedRow.Cells["Hatalı Miktar"].Value?.ToString();
            var adet = selectedRow.Cells["Adet"].Value?.ToString();
            var toplamMiktar = selectedRow.Cells["Toplam Miktar"].Value?.ToString();
            var tarih = selectedRow.Cells["Tarih"].Value?.ToString();
            var kayipZaman = selectedRow.Cells["Kayıp Zaman"].Value?.ToString();
            var zamanTuru = selectedRow.Cells["Zaman Türü"].Value?.ToString();
            var hataTipi = selectedRow.Cells["Hata Tipi"].Value?.ToString();
            var aciklama = selectedRow.Cells["Aciklama"].Value?.ToString();
            var tedarikci = selectedRow.Cells["Tedarikçi"].Value?.ToString();
            var ozet = selectedRow.Cells["Ozet"].Value?.ToString();
            var hataBolumu = selectedRow.Cells["Hata Bolumu"].Value?.ToString();
            var raporuHazirlayan = selectedRow.Cells["Raporu Hazirlayan"].Value?.ToString();
            var hatayiBulanBirim = selectedRow.Cells["Hatayı Bulan Birim"].Value?.ToString();
            var kokNeden = selectedRow.Cells["Kök Neden"].Value?.ToString();
            var aksiyon = selectedRow.Cells["Aksiyon"].Value?.ToString();
            var sonuc = selectedRow.Cells["Sonuç"].Value?.ToString();
            var degerlendiren = selectedRow.Cells["Değerlendiren"].Value?.ToString();
            var kokNedenAksiyon = selectedRow.Cells["Kök Neden Aksiyon"].Value?.ToString();
            var resim = selectedRow.Cells["Resim"].Value?.ToString();
            var kapanisTarihi = selectedRow.Cells["Kapanış Tarihi"].Value?.ToString();
            var terminTarihi = selectedRow.Cells["Termin Tarihi"].Value?.ToString();
            var urunTipi = selectedRow.Cells["Ürün Tipi"].Value?.ToString();
            var duzelticiFaaliyetVarMi = selectedRow.Cells["Düzeltici Faaliyet Var Mı?"].Value?.ToString();
            var durum = selectedRow.Cells["Durum"].Value?.ToString();

            // Excel işlemleri
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                // Excel dosyasının yolu
                string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Dosyalar", "UygunOlmayanUrunRaporu.xlsx");

                if (!File.Exists(excelFilePath))
                {
                    MessageBox.Show("Excel dosyası bulunamadı!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets["FR.87.01 (R1)"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Çalışma sayfası bulunamadı!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Ürün tipi işaretleme
                    if (urunTipi == "Ham Malzeme") worksheet.Cells["A9"].Value = "X";
                    else if (urunTipi == "Standart Malzeme") worksheet.Cells["D9"].Value = "X";
                    else if (urunTipi == "Yarı Mamul") worksheet.Cells["H9"].Value = "X";
                    else if (urunTipi == "Bitmiş Ürün") worksheet.Cells["K9"].Value = "X";

                    // Verileri Excel'e yaz
                    worksheet.Cells["A4"].Value = $"Ürün Kodu: {urunKodu}";
                    worksheet.Cells["E4"].Value = $"Ürün Adı: {urunAdi}";
                    worksheet.Cells["J4"].Value = $"Sipariş No/Proje Kodu: {siparisNo}";
                    worksheet.Cells["A5"].Value = $"Hatalı Miktar: {hataliMiktar}";
                    worksheet.Cells["E5"].Value = $"Toplam Miktar: {toplamMiktar}";
                    worksheet.Cells["K5"].Value = $"İşlem Kayıp Zamanı: {kayipZaman}";
                    worksheet.Cells["A6"].Value = $"Tespit Eden Bölüm: {hatayiBulanBirim}";
                    worksheet.Cells["A7"].Value = $"Dış Tedarikçi veya Uygun Olmayan Ürünün Oluştuğu Bölüm: {hataBolumu}";
                    worksheet.Cells["A25"].Value = $"Raporu Hazırlayan: {raporuHazirlayan}";

                    // Açıklamalar için hücreleri birleştir ve metni ayarla
                    worksheet.Cells["A16"].Value = $"Hatanın Tanımı: {aciklama}";
                    worksheet.Cells["A21"].Value = $"Hatanın Kök Nedeni: {kokNeden}";
                    worksheet.Cells["A29"].Value = $"DEĞERLENDİRME - YAPILACAK FAALİYETLER: {aksiyon}";
                    worksheet.Cells["A41"].Value = $"DEĞERLENDİRME NOTLARI / SONUÇ: {sonuc}";
                    worksheet.Cells["A46"].Value = $"Düzeltici Faaliyet Gerekiyor: {duzelticiFaaliyetVarMi}";

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
        // Hücre birleştirme ve metin ekleme fonksiyonu
        private void MergeAndSetCell(ExcelWorksheet worksheet, string range, string text)
        {
            worksheet.Cells[range].Merge = true;
            worksheet.Cells[range].Style.WrapText = true;
            worksheet.Cells[range].Value = text;
        }

        private void gERİDÖNToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UygunOlmayanDurum UGOD = new UygunOlmayanDurum();
            UGOD.Show();
            this.Close();
        }



        public void filtreleme()
        {
       
        }
        private void advancedDataGridView1_FilterStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.FilterEventArgs e)
        {
            table.DefaultView.RowFilter = advancedDataGridView1.FilterString;
            RenklendirDurum(); // Filtre değiştikçe renkleri tazele
        }

        private void advancedDataGridView1_SortStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.SortEventArgs e)
        {
            table.DefaultView.Sort = advancedDataGridView1.SortString;
        }

        private void advancedDataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            RenklendirDurum();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
                dbContext?.Dispose();
                UOD?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
