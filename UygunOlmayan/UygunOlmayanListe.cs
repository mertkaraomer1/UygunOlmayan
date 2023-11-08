using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System.Data;
using UygunOlmayan.MyDb;

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

        private void button1_Click(object sender, EventArgs e)
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
        private void UygunOlmayanListe_Load(object sender, EventArgs e)
        {
            // Sütunları ekle
            table.Columns.Add("Urun Id");
            table.Columns.Add("Urun Kodu");
            table.Columns.Add("Urun Adi");
            table.Columns.Add("Siparis No");
            table.Columns.Add("Hatalı Miktar");
            table.Columns.Add("Tarih");
            table.Columns.Add("Kayıp Zaman");
            table.Columns.Add("HataTipi");
            table.Columns.Add("Aciklama");
            table.Columns.Add("Ozet");
            table.Columns.Add("Hata Bolumu");
            table.Columns.Add("Raporu Hazirlayan");
            table.Columns.Add("Resim");
            table.Columns.Add("Hatayı Bulan Birim");
            table.Columns.Add("Kök Neden");
            table.Columns.Add("Aksiyon");
            table.Columns.Add("Sonuc");


            var hataliUrun = dbContext.hataliUruns.ToList();

            // İhtiyacınıza göre daha fazla sütun ekleyebilirsiniz.

            // Verileri eklemek için bir döngü kullanabilirsiniz
            foreach (var urun in hataliUrun)
            {
                table.Rows.Add(urun.UrunId, urun.UrunKodu, urun.UrunAdi, urun.SiparisNo, urun.HatalıMiktar, urun.Tarih, urun.KayıpZaman, urun.HataTipi, urun.Aciklama, urun.Ozet, urun.HataBolumu, urun.RaporuHazirlayan, urun.Resim, urun.HatayıBulanBirim, urun.KokNeden, urun.Aksiyon, urun.Sonuc);
                // İhtiyacınıza göre daha fazla sütun eklerseniz, burada da verileri ilgili sütunlara eklemeniz gerekir.
            }

            // DataTable'ı DataGridView'e bağlayın
            advancedDataGridView1.DataSource = table;
        }
        Resim RS;
        private void advancedDataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 1)
            {

                Resim RS = new Resim();

                DataGridViewCell clickedCell = advancedDataGridView1.Rows[e.RowIndex].Cells[12]; // Tıklanan hücreyi al
                string cellValue = clickedCell.Value.ToString();
                RS.picturebox132 = cellValue;
                RS.Show();

            }
        }
    }
}
