using System.Data;
using System.Net.Mail;
using System.Net;
using System.Windows.Forms;
using UygunOlmayan.MyDb;
using UygunOlmayan.Tables;
using OfficeOpenXml;
using System.Diagnostics;

namespace UygunOlmayan
{
    public partial class UygunOlmayanListe : Form
    {
        MyDbContext dbContext;
        public UygunOlmayanListe()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }


        DataTable table = new DataTable();
        public void UygunOlmayanVol1()
        {
            // Önce tabloyu temizle
            table.Clear();
            table.Columns.Clear();
            advancedDataGridView1.Columns.Clear();

            // Sütunları ekle
            string[] columnNames = {
        "Urun Id", "Urun Kodu", "Urun Adi", "Siparis No", "Hatalı Miktar", "Adet",
        "Toplam Miktar", "Tarih", "Kayıp Zaman", "Zaman Türü", "Hata Tipi", "Aciklama",
        "Tedarikçi", "Ozet", "Hata Bolumu", "Raporu Hazirlayan", "Hatayı Bulan Birim",
        "Kök Neden", "Aksiyon", "Sonuç", "Değerlendiren", "Kök Neden Aksiyon", "Resim",
        "Kapanış Tarihi", "Termin Tarihi", "Ürün Tipi", "Düzeltici Faaliyet Var Mı?", "Durum","Aksiyon Alancak Bölüm"
    };

            foreach (string columnName in columnNames)
            {
                table.Columns.Add(columnName);
            }

            var hataliUrunList = dbContext.hataliUruns
                .Where(x => x.urunimza == new Guid())
                .ToList();




            foreach (var urun in hataliUrunList)
            {
                var aksiyonbölümü = dbContext.aksiyonAlacakBölüms
                    .Where(x => x.UygunOlmayanId == urun.UrunId)
                    .OrderByDescending(x => x.Tarihi) // Tarihi en yeni olanı seç
                    .FirstOrDefault();

                string durum = urun.Durum == "False" ? "İŞLEMİ DEVAM EDİYOR." : "DEĞERLENDİRMEYİ BEKLİYOR.";
                string aksiyonAlındımı = urun.AksiyonAlındı == "False" ? durum : "AKSİYON ALINDI";
                durum = aksiyonAlındımı;
                table.Rows.Add(
                    urun.UrunId, urun.UrunKodu, urun.UrunAdi, urun.SiparisNo, urun.HatalıMiktar,
                    urun.Adet, urun.toplamMiktar, urun.Tarih.ToString("yyyy.MM.dd"), urun.KayıpZaman,
                    urun.ZamanCinsi, urun.HataTipi, urun.Aciklama, urun.Tedarikci, urun.Ozet,
                    urun.HataBolumu, urun.RaporuHazirlayan, urun.HatayıBulanBirim, urun.KokNeden,
                    urun.Aksiyon, urun.Sonuc, urun.Degerlendiren, urun.KokNedenAksiyon,
                    urun.Resim, urun.KapanısTarihi, urun.TerminTarihi, urun.uruntipi,
                    urun.DuzelticiFaliyetDurum, durum, aksiyonbölümü?.AksiyonBölümü
                );
            }

            // DataTable'ı DataGridView'e bağla
            advancedDataGridView1.DataSource = table;
            // DataGridView'e imza sütunu ekle
            DataGridViewImageColumn buttonColumn1 = new DataGridViewImageColumn
            {
                HeaderText = "Aksiyon Alındı Butonu",
                Image = Image.FromFile("delete.png"),
                ImageLayout = DataGridViewImageCellLayout.Zoom
            };
            advancedDataGridView1.Columns.Add(buttonColumn1);
            // DataGridView'e imza sütunu ekle
            DataGridViewImageColumn buttonColumn2 = new DataGridViewImageColumn
            {
                HeaderText = "Yetkili İmza Alanı",
                Image = Image.FromFile("imza.png"),
                ImageLayout = DataGridViewImageCellLayout.Zoom
            };
            advancedDataGridView1.Columns.Add(buttonColumn2);


            // **Durum sütunu için renklendirme işlemi**
            for (int i = 0; i < advancedDataGridView1.Rows.Count; i++)
            {
                DataGridViewRow row = advancedDataGridView1.Rows[i];

                // "Durum" sütununun index'ini al
                int durumColumnIndex = advancedDataGridView1.Columns["Durum"].Index;

                if (row.Cells[durumColumnIndex].Value != null)
                {
                    string durumText = row.Cells[durumColumnIndex].Value.ToString();

                    if (durumText == "İŞLEMİ DEVAM EDİYOR.")
                    {
                        row.Cells[durumColumnIndex].Style.BackColor = Color.Yellow;
                        row.Cells[durumColumnIndex].Style.ForeColor = Color.Black;
                    }
                    else if (durumText == "DEĞERLENDİRMEYİ BEKLİYOR.")
                    {
                        row.Cells[durumColumnIndex].Style.BackColor = Color.Red;
                        row.Cells[durumColumnIndex].Style.ForeColor = Color.White;
                    }
                    else if (durumText == "AKSİYON ALINDI")
                    {
                        row.Cells[durumColumnIndex].Style.BackColor = Color.Blue;
                        row.Cells[durumColumnIndex].Style.ForeColor = Color.White;
                    }
                }
            }
        }

        private void UygunOlmayanListe_Load(object sender, EventArgs e)
        {
            UygunOlmayanVol1();
        }

        private UygunOlmayanDurum UOD; // Formun bir örneği için alan
        private void advancedDataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 1)
            {
                // Uygun Olmayan Durum formunun mevcut olup olmadığını kontrol et
                if (UOD == null || UOD.IsDisposed)
                {
                    UOD = new UygunOlmayanDurum();
                }
                else
                {
                    // Form zaten açıksa, mevcut formu güncelleyelim
                    UOD.BringToFront(); // Formu ön plana getir
                    return; // Fonksiyondan çık
                }

                // Hücre değerlerini alma
                string UrunID = advancedDataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                string UrunKodu = advancedDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                string UrunAdi = advancedDataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                string SipNo = advancedDataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                string HataMik = advancedDataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                string toplamMik = advancedDataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
                string KayıpZamna = advancedDataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
                string HataTipi = advancedDataGridView1.Rows[e.RowIndex].Cells[10].Value.ToString();
                string Aciklama = advancedDataGridView1.Rows[e.RowIndex].Cells[11].Value.ToString();
                string Tedarikci = advancedDataGridView1.Rows[e.RowIndex].Cells[12].Value.ToString();
                string Ozet = advancedDataGridView1.Rows[e.RowIndex].Cells[13].Value.ToString();
                string HataBolumu = advancedDataGridView1.Rows[e.RowIndex].Cells[14].Value.ToString();
                string RHazırlayan = advancedDataGridView1.Rows[e.RowIndex].Cells[15].Value.ToString();
                string HatayıBulanBolum = advancedDataGridView1.Rows[e.RowIndex].Cells[16].Value.ToString();
                string KokNeden = advancedDataGridView1.Rows[e.RowIndex].Cells[17].Value.ToString();
                string Aksiyon = advancedDataGridView1.Rows[e.RowIndex].Cells[18].Value.ToString();
                string Sonuc = advancedDataGridView1.Rows[e.RowIndex].Cells[19].Value.ToString();
                string Degerlendiren = advancedDataGridView1.Rows[e.RowIndex].Cells[20].Value.ToString();
                string KokNedenAksiyon = advancedDataGridView1.Rows[e.RowIndex].Cells[21].Value.ToString();
                string Resim = advancedDataGridView1.Rows[e.RowIndex].Cells[22].Value.ToString();
                string terminTarihiString = advancedDataGridView1.Rows[e.RowIndex].Cells[24].Value.ToString();
                string uruntipi2 = advancedDataGridView1.Rows[e.RowIndex].Cells[25].Value.ToString();
                var resim = dbContext.hataliUruns.Where(x => x.UrunId == Convert.ToInt32(UrunID)).Select(x => x.Resim).FirstOrDefault();
                string DuzelticiFaliyetDurum2 = advancedDataGridView1.Rows[e.RowIndex].Cells[26].Value.ToString();

                // DateTime'a dönüştür
                if (DateTime.TryParse(terminTarihiString, out DateTime terminTarihi))
                {
                    UOD.TerminTarihi = terminTarihi;
                }
                else
                {
                    MessageBox.Show("Termin tarihi geçersiz!");
                }

                // Formu doldurma
                UOD.UrunID1 = UrunID;
                UOD.UrunKodu1 = UrunKodu;
                UOD.UrunAdi1 = UrunAdi;
                UOD.SipNo1 = SipNo;
                UOD.HataMik1 = HataMik;
                UOD.toplamMik1 = toplamMik;
                UOD.KayıpZaman1 = KayıpZamna;
                UOD.Hatatipi1 = HataTipi;
                UOD.Acıklama1 = Aciklama;
                UOD.ozet1 = Ozet;
                UOD.HataBolumu1 = HataBolumu;
                UOD.Rhazirlayan1 = RHazırlayan;
                UOD.HataBulanBlum1 = HatayıBulanBolum;
                UOD.KokNeden1 = KokNeden;
                UOD.Aksiyon1 = Aksiyon;
                UOD.Sonuc1 = Sonuc;
                UOD.tedarikci = Tedarikci;
                UOD.Degerlendiren = Degerlendiren;
                UOD.KokNedenAksiyon = KokNedenAksiyon;
                UOD.LoadImageFromBytes(resim);
                UOD.uruntipi1 = uruntipi2;
                UOD.DuzelticiFaliyetDurum1 = DuzelticiFaliyetDurum2;
                UOD.ButtonGuncelle = true;
               UOD.ButtonKaydet = false;
                // Formu göster
                UOD.Show();
            }
        }

        private void advancedDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 30)
            {
                // Seçilen satırın verilerini almak için DataGridView'den erişin
                DataGridViewRow selectedRow = advancedDataGridView1.Rows[e.RowIndex];


                string UrunKod = selectedRow.Cells["Urun Kodu"].Value.ToString();
                string SipNo = selectedRow.Cells["Siparis No"].Value.ToString();


                // UrunKodu ve UrunAdi'ye göre eşleşen HataliUrun nesnesini bulun
                var eskiHataliUrun = dbContext.hataliUruns
                    .FirstOrDefault(u => u.UrunKodu == UrunKod && u.SiparisNo == SipNo);

                // Eğer eşleşen bir nesne bulunduysa, alanları güncelleyin
                if (eskiHataliUrun != null)
                {
                    eskiHataliUrun.urunimza = Guid.NewGuid();
                    eskiHataliUrun.KapanısTarihi = DateTime.Now.Date;
                    // Değişiklikleri veritabanına kaydedin
                    dbContext.SaveChanges();
                    MessageBox.Show("Güncellendi.");
                    UygunOlmayanVol1();
                }
            }
            else if (e.RowIndex >= 0 && e.ColumnIndex == 29)
            {
                // Seçilen satırın verilerini almak için DataGridView'den erişin
                DataGridViewRow selectedRow = advancedDataGridView1.Rows[e.RowIndex];


                string UrunKod = selectedRow.Cells["Urun Kodu"].Value.ToString();
                string SipNo = selectedRow.Cells["Siparis No"].Value.ToString();


                // UrunKodu ve UrunAdi'ye göre eşleşen HataliUrun nesnesini bulun
                var eskiHataliUrun = dbContext.hataliUruns
                    .FirstOrDefault(u => u.UrunKodu == UrunKod && u.SiparisNo == SipNo);

                // Eğer eşleşen bir nesne bulunduysa, alanları güncelleyin
                if (eskiHataliUrun != null)
                {
                    eskiHataliUrun.AksiyonAlındı = "True";

                    dbContext.SaveChanges();
                    MessageBox.Show("Güncellendi.");
                    UygunOlmayanVol1();
                }
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // DataGridView'deki verileri bir DataTable'a kopyalayın
            DataTable dt = new DataTable();

            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
                // Eğer ValueType null ise, varsayılan bir veri türü kullanabilirsiniz.
                Type columnType = column.ValueType ?? typeof(string);
                dt.Columns.Add(column.HeaderText, columnType);
            }

            // Satırları ekle
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                DataRow dataRow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dataRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dataRow);
            }

            // Excel uygulamasını başlatın
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            // Yeni bir Excel çalışma kitabı oluşturun
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            // DataTable'ı Excel çalışma sayfasına aktarın (tablo başlıklarını da dahil etmek için)
            int rowIndex = 1;

            // Başlıkları yaz
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                worksheet.Cells[1, j + 1] = dt.Columns[j].ColumnName;
                worksheet.Cells[1, j + 1].Font.Bold = true;
            }

            // Verileri yaz
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                rowIndex++;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    worksheet.Cells[rowIndex, j + 1] = dt.Rows[i][j].ToString();
                }
            }
        }


        private void lİSTEYİEXCELEAKTARToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // DataGridView'deki verileri bir DataTable'a kopyalayın
            DataTable dt = new DataTable();

            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
                // Eğer ValueType null ise, varsayılan bir veri türü kullanabilirsiniz.
                Type columnType = column.ValueType ?? typeof(string);
                dt.Columns.Add(column.HeaderText, columnType);
            }

            // Satırları ekle
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                DataRow dataRow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dataRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dataRow);
            }

            // Excel uygulamasını başlatın
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            // Yeni bir Excel çalışma kitabı oluşturun
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            // DataTable'ı Excel çalışma sayfasına aktarın (tablo başlıklarını da dahil etmek için)
            int rowIndex = 1;

            // Başlıkları yaz
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                worksheet.Cells[1, j + 1] = dt.Columns[j].ColumnName;
                worksheet.Cells[1, j + 1].Font.Bold = true;
            }

            // Verileri yaz
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                rowIndex++;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    worksheet.Cells[rowIndex, j + 1] = dt.Rows[i][j].ToString();
                }
            }
        }

        private void ePOSTAGÖNDERToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewRow selectedRow = advancedDataGridView1.CurrentRow;
            if (selectedRow != null)
            {
                var urunIdValue = selectedRow.Cells["Urun Id"].Value?.ToString(); // "UrunId" sütunundaki değeri string olarak al
                var HataBolumu = selectedRow.Cells["Hata Bolumu"].Value?.ToString();
                string subject = "UYGUN OLMAYAN ÜRÜN KONTROL FORMU";
                string urunId = urunIdValue; // urunId'yi buraya ata
                string body;

                // Bölümlere göre birden fazla e-posta adresi içeren sözlük tanımlandı.
                var emailAddresses = new Dictionary<string, List<string>>
                {
                    { "Montaj", new List<string> { "dturkan@icmmakina.com", "oocak@icmmakina.com" } },
                    { "Tasarım", new List<string> { "mbayram@icmmakina.com", "uulusoy@icmmakina.com", "ecobanbas@icmmakina.com", "etaskin@icmmakina.com", "ucsari@icmmakina.com", "minan@icmmakina.com", "akucukler@icmmakina.com", "morhan@icmmakina.com" } },
                    { "İmalat", new List<string> { "pyesilyurt@icmmakina.com", "skoca@icmmakina.com", "hetanta@icmmakina.com" } },
                    { "Otomasyon", new List<string> { "otomasyon.proje@icmmakina.com", "tozpinar@icmmakina.com", "egozluk@icmmakina.com", "bguden@icmmakina.com" , "byanik@icmmakina.com" } },
                    { "Satınalma", new List<string> { "satinalma@icmmakina.com" } },
                    { "Planlama", new List<string> { "shaci@icmmakina.com", "sbuyukay@icmmakina.com" } },
                    { "Kalite Kontrol", new List<string> { "oarslan@icmmakina.com" } },
                    { "Satış Sonrası", new List<string> { "hsokmen@icmmakina.com", "dtacyildiz@icmmakina.com" } },
                    { "Muhasebe", new List<string> { "bozcan@icmmakina.com", "mcelik@icmmakina.com" } },
                    { "Fabrika Müdürü", new List<string> { "ddeniz@icmmakina.com" } }
                };

                // Seçilen bölüme göre e-posta adresleri belirleniyor.
                if (emailAddresses.TryGetValue(HataBolumu, out List<string> emailList))
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
            else
            {
                MessageBox.Show("Lütfen bir satır seçin!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

    }
}
