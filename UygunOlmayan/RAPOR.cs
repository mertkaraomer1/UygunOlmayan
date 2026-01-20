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
using System.IO;
using System.Diagnostics;
using OfficeOpenXml;
using Microsoft.EntityFrameworkCore;

namespace UygunOlmayan
{
    public partial class RAPOR : Form
    {
        MyDbContext dbContext;
        public RAPOR()
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

        private async void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedItem == null) return;
            string secilenBolum = comboBox3.SelectedItem.ToString();

            this.Cursor = Cursors.WaitCursor; // Kullanıcıya işlemin devam ettiğini bildir

            try
            {
                table.Clear();
                if (table.Columns.Count == 0)
                {
                    table.Columns.Add("Sıra No", typeof(int));
                    table.Columns.Add("Ürün Kodu", typeof(string));
                    table.Columns.Add("Sipariş No", typeof(string));
                    table.Columns.Add("Hata Tipi", typeof(string));
                    table.Columns.Add("Hata Bölümü", typeof(string));
                    table.Columns.Add("Kök Neden", typeof(string));
                    table.Columns.Add("Kök Neden Aksiyon", typeof(string));
                    table.Columns.Add("Açıklama", typeof(string));
                    table.Columns.Add("Tedarikçi", typeof(string));
                    table.Columns.Add("Kayıp Gün", typeof(double));
                    table.Columns.Add("Tarih", typeof(string));
                    table.Columns.Add("Kapanış Tarihi", typeof(string));
                }

                // AsNoTracking() ile veri çekme hızını artırıyoruz (Sadece okuma amaçlı olduğu için)
                var query = dbContext.hataliUruns
                    .AsNoTracking()
                    .Where(x => x.urunimza != Guid.Empty);

                if (secilenBolum != "TÜMÜ")
                {
                    query = query.Where(x => x.HataBolumu == secilenBolum);
                }

                var hataliUrunList = await query.Select(x => new
                {
                    x.HataTipi,
                    x.HataBolumu,
                    x.SiparisNo,
                    x.UrunKodu,
                    x.KokNeden,
                    x.KokNedenAksiyon,
                    x.Aciklama,
                    x.Tedarikci,
                    x.Tarih,
                    x.KapanısTarihi,
                    kayıpGun = Math.Round((x.KapanısTarihi - x.Tarih).TotalDays, 0)
                }).ToListAsync();

                int sayac = 1;
                foreach (var urun in hataliUrunList)
                {
                    table.Rows.Add(
                        sayac++,
                        urun.UrunKodu ?? "",
                        urun.SiparisNo ?? "",
                        urun.HataTipi ?? "",
                        urun.HataBolumu ?? "",
                        urun.KokNeden ?? "",
                        urun.KokNedenAksiyon ?? "",
                        urun.Aciklama ?? "",
                        urun.Tedarikci ?? "",
                        urun.kayıpGun,
                        urun.Tarih.ToString("yyyy.MM.dd"),
                        urun.KapanısTarihi.ToString("yyyy.MM.dd")
                    );
                }

                advancedDataGridView1.DataSource = table;

                // Performans için otomatik boyutlandırmayı bir kez yapıyoruz
                advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                advancedDataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Rapor yüklenirken hata oluştu: {ex.Message}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default; // İmleci normale döndür
            }
        }

        private void lİSTEYİEXCELEAKTARToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("ANALİZ RAPORU");

                for (int i = 0; i < advancedDataGridView1.Columns.Count; i++)
                    worksheet.Cells[1, i + 1].Value = advancedDataGridView1.Columns[i].HeaderText;

                for (int i = 0; i < advancedDataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < advancedDataGridView1.Columns.Count; j++)
                        worksheet.Cells[i + 2, j + 1].Value = advancedDataGridView1.Rows[i].Cells[j].Value;
                }

                string path = Path.Combine(Path.GetTempPath(), $"AnalizRaporu_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                File.WriteAllBytes(path, package.GetAsByteArray());
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
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
