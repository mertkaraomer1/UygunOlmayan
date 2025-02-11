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

namespace UygunOlmayan
{
    public partial class UygunOlmayanRapor : Form
    {
        MyDbContext dbContext;
        public UygunOlmayanRapor()
        { 
            dbContext = new MyDbContext();
            InitializeComponent();
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
        "Kapanış Tarihi", "Termin Tarihi", "Ürün Tipi", "Düzeltici Faaliyet Var Mı?", "İmza"
    };

            foreach (string columnName in columnNames)
            {
                table.Columns.Add(columnName);
            }

            // Verileri çek
            var hataliUrunList = dbContext.hataliUruns.Where(x => x.urunimza != new Guid()).ToList();

            foreach (var urun in hataliUrunList)
            {
                string durum = urun.Durum == "True" ? "İŞLEMİ DEVAM EDİYOR." : "DEĞERLENDİRMEYİ BEKLİYOR.";

                table.Rows.Add(
                    urun.UrunId, urun.UrunKodu, urun.UrunAdi, urun.SiparisNo, urun.HatalıMiktar,
                    urun.Adet, urun.toplamMiktar, urun.Tarih.ToString("yyyy.MM.dd"), urun.KayıpZaman,
                    urun.ZamanCinsi, urun.HataTipi, urun.Aciklama, urun.Tedarikci, urun.Ozet,
                    urun.HataBolumu, urun.RaporuHazirlayan, urun.HatayıBulanBirim, urun.KokNeden,
                    urun.Aksiyon, urun.Sonuc, urun.Degerlendiren, urun.KokNedenAksiyon,
                    urun.Resim, urun.KapanısTarihi, urun.TerminTarihi, urun.uruntipi,
                    urun.DuzelticiFaliyetDurum,urun.urunimza
                );
            }

            // DataTable'ı DataGridView'e bağla
            advancedDataGridView1.DataSource = table;
        }
        private void UygunOlmayanRapor_Load(object sender, EventArgs e)
        {
            UygunOlmayanVol1();
        }
    }
}
