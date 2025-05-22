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
    public partial class RAPOR : Form
    {
        MyDbContext dbContext;
        public RAPOR()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }
        DataTable table = new DataTable();

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            string HatanınOlustuguBolum = comboBox3.SelectedItem.ToString();
            table.Clear(); // tabloyu temizle
            table.Columns.Clear(); // kolonları temizle
            table.Rows.Clear(); // satırları temizle
            table.Columns.Add("Satır Sayısı");
            table.Columns.Add("UrunKodu");
            table.Columns.Add("SiparisNo");
            table.Columns.Add("HataTipi");
            table.Columns.Add("HataBolumu");
            table.Columns.Add("KokNeden");
            table.Columns.Add("KokNedenAksiyon");
            table.Columns.Add("Açıklama"); // KokNedenAksiyon kolonu da şart!
            table.Columns.Add("Tedarikçi"); // KokNedenAksiyon kolonu da şart!
            table.Columns.Add("KayıpGun");  // kayıp gün kolonu da şart!
            table.Columns.Add("Tarih"); // Tarih kolonu da şart!
            table.Columns.Add("KapanisTarihi"); // Kapanış tarihi kolonu da şart!

            var hataliUrunList = dbContext.hataliUruns
                .Where(x => x.HataBolumu == HatanınOlustuguBolum && x.urunimza != new Guid())
                .Select(x => new
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
                    kayıpGun = Math.Round((x.KapanısTarihi - x.Tarih).TotalDays, 0) // sadece gün sayısını verir
                })
                .ToList();
            int sayac = 1;
            foreach (var urun in hataliUrunList)
            {
                table.Rows.Add(sayac++, urun.UrunKodu, urun.SiparisNo, urun.HataTipi, urun.HataBolumu, urun.KokNeden, urun.KokNedenAksiyon,urun.Aciklama,urun.Tedarikci, urun.kayıpGun, urun.Tarih, urun.KapanısTarihi);
            }
            advancedDataGridView1.DataSource = table;
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
    }
}
