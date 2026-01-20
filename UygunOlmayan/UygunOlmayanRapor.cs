using Microsoft.EntityFrameworkCore;
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
using System.IO;
using System.Diagnostics;
using OfficeOpenXml;

namespace UygunOlmayan
{
    public partial class UygunOlmayanRapor : Form
    {
        MyDbContext dbContext;
        public UygunOlmayanRapor()
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

        private async void toolStripButton1_Click(object sender, EventArgs e)
        {
            await UygunOlmayanVol1Async();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ExportToExcelFast();
        }

        private void ExportToExcelFast()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("RAPOR");

                for (int i = 0; i < advancedDataGridView1.Columns.Count; i++)
                    worksheet.Cells[1, i + 1].Value = advancedDataGridView1.Columns[i].HeaderText;

                for (int i = 0; i < advancedDataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < advancedDataGridView1.Columns.Count; j++)
                        worksheet.Cells[i + 2, j + 1].Value = advancedDataGridView1.Rows[i].Cells[j].Value;
                }

                string path = Path.Combine(Path.GetTempPath(), $"UygunOlmayanRapor_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                File.WriteAllBytes(path, package.GetAsByteArray());
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
        }
        DataTable table = new DataTable();
        public async Task UygunOlmayanVol1Async()
        {
            this.Cursor = Cursors.WaitCursor;
            try
            {
                table.Clear();
                if (table.Columns.Count == 0)
                {
                    string[] columnNames = {
                        "Urun Id", "Urun Kodu", "Urun Adi", "Siparis No", "Hatalı Miktar", "Adet",
                        "Toplam Miktar", "Tarih", "Kayıp Zaman", "Zaman Türü", "Hata Tipi", "Aciklama",
                        "Tedarikçi", "Ozet", "Hata Bolumu", "Raporu Hazirlayan", "Hatayı Bulan Birim",
                        "Kök Neden", "Aksiyon", "Sonuç", "Değerlendiren", "Kök Neden Aksiyon", "Resim Durumu",
                        "Kapanış Tarihi", "Termin Tarihi", "Ürün Tipi", "Düzeltici Faaliyet Var Mı?", "İmza Durumu"
                    };
                    foreach (string columnName in columnNames)
                        table.Columns.Add(columnName);
                }

                // Projeksiyon ile sessizce sadece gerekli verileri çekiyoruz (Resim BLOB'unu çekmiyoruz)
                var hataliUrunList = await dbContext.hataliUruns
                    .AsNoTracking()
                    .Where(x => x.urunimza != Guid.Empty)
                    .Select(u => new
                    {
                        u.UrunId, u.UrunKodu, u.UrunAdi, u.SiparisNo, u.HatalıMiktar, u.Adet,
                        u.toplamMiktar, u.Tarih, u.KayıpZaman, u.ZamanCinsi, u.HataTipi, u.Aciklama,
                        u.Tedarikci, u.Ozet, u.HataBolumu, u.RaporuHazirlayan, u.HatayıBulanBirim,
                        u.KokNeden, u.Aksiyon, u.Sonuc, u.Degerlendiren, u.KokNedenAksiyon,
                        ResimVar = u.Resim != null ? "VAR" : "YOK",
                        u.KapanısTarihi, u.TerminTarihi, u.uruntipi, u.DuzelticiFaliyetDurum,
                        u.urunimza
                    }).ToListAsync();

                foreach (var urun in hataliUrunList)
                {
                    table.Rows.Add(
                        urun.UrunId, urun.UrunKodu, urun.UrunAdi, urun.SiparisNo, urun.HatalıMiktar,
                        urun.Adet, urun.toplamMiktar, urun.Tarih.ToString("yyyy.MM.dd"), urun.KayıpZaman,
                        urun.ZamanCinsi, urun.HataTipi, urun.Aciklama, urun.Tedarikci, urun.Ozet,
                        urun.HataBolumu, urun.RaporuHazirlayan, urun.HatayıBulanBirim, urun.KokNeden,
                        urun.Aksiyon, urun.Sonuc, urun.Degerlendiren, urun.KokNedenAksiyon,
                        urun.ResimVar, urun.KapanısTarihi.ToString("yyyy.MM.dd"), 
                        urun.TerminTarihi?.ToString("yyyy.MM.dd") ?? "", urun.uruntipi,
                        urun.DuzelticiFaliyetDurum, urun.urunimza == Guid.Empty ? "İMZASIZ" : "İMZALI"
                    );
                }

                advancedDataGridView1.DataSource = table;
                advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                advancedDataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Veri yükleme hatası: {ex.Message}");
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
        private async void UygunOlmayanRapor_Load(object sender, EventArgs e)
        {
            await UygunOlmayanVol1Async();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
                dbContext?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
