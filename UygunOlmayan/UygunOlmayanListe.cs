using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System.Data;
using System.Windows.Forms;
using UygunOlmayan.MyDb;
using UygunOlmayan.Tables;
using Zuby.ADGV;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

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
        public void Uygunolmayanvol1()
        {

            table.Columns.Clear();
            table.Rows.Clear();
            advancedDataGridView1.Columns.Clear();
            // Sütunları ekle
            table.Columns.Add("Urun Id");
            table.Columns.Add("Urun Kodu");
            table.Columns.Add("Urun Adi");
            table.Columns.Add("Siparis No");
            table.Columns.Add("Hatalı Miktar");
            table.Columns.Add("Adet");
            table.Columns.Add("Tarih");
            table.Columns.Add("Kayıp Zaman");
            table.Columns.Add("Zaman Türü");
            table.Columns.Add("HataTipi");
            table.Columns.Add("Aciklama");
            table.Columns.Add("Tedarikçi");
            table.Columns.Add("Ozet");
            table.Columns.Add("Hata Bolumu");
            table.Columns.Add("Raporu Hazirlayan");
            table.Columns.Add("Resim");
            table.Columns.Add("Hatayı Bulan Birim");
            table.Columns.Add("Kök Neden");
            table.Columns.Add("Aksiyon");
            table.Columns.Add("Sonuc");




            var hataliUrun = dbContext.hataliUruns
                .ToList();


            // İhtiyacınıza göre daha fazla sütun ekleyebilirsiniz.

            foreach (var urun in hataliUrun)
            {
                table.Rows.Add(urun.UrunId, urun.UrunKodu, urun.UrunAdi, urun.SiparisNo, urun.HatalıMiktar, urun.Adet, urun.Tarih.ToString("yyyy.MM.dd"), urun.KayıpZaman, urun.ZamanCinsi, urun.HataTipi, urun.Aciklama, urun.Ozet, urun.HataBolumu, urun.RaporuHazirlayan, urun.Resim, urun.HatayıBulanBirim, urun.KokNeden, urun.Aksiyon, urun.Sonuc);
            }

            // DataTable'ı DataGridView'e bağlayın
            advancedDataGridView1.DataSource = table;
            // DataGridView kontrolünüze bir buton sütunu ekleyin.
            DataGridViewImageColumn buttonColumn = new DataGridViewImageColumn();
            buttonColumn.HeaderText = "Durumu Kapat"; // Sütun başlığı
            buttonColumn.Image = Image.FromFile("delete.png"); // Silme resmini belirtin
            buttonColumn.ImageLayout = DataGridViewImageCellLayout.Zoom; // Resmi düzgün görüntülemek için ayar
            advancedDataGridView1.Columns.Add(buttonColumn);


        }
        private void UygunOlmayanListe_Load(object sender, EventArgs e)
        {
            Uygunolmayanvol1();
        }

        UygunOlmayanDurum UOD;
        private void advancedDataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 1)
            {
                if (UOD == null || UOD.IsDisposed)
                {
                    UOD = new UygunOlmayanDurum();
                    string UrunKodu = advancedDataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    string UrunAdi = advancedDataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    string SipNo = advancedDataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    string HataMik = advancedDataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                    string KayıpZamna = advancedDataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
                    string HataTipi = advancedDataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString();
                    string Aciklama = advancedDataGridView1.Rows[e.RowIndex].Cells[10].Value.ToString();
                    string Tedarikci = advancedDataGridView1.Rows[e.RowIndex].Cells[11].Value.ToString();
                    string Ozet = advancedDataGridView1.Rows[e.RowIndex].Cells[12].Value.ToString();
                    string HataBolumu = advancedDataGridView1.Rows[e.RowIndex].Cells[13].Value.ToString();
                    string RHazırlayan = advancedDataGridView1.Rows[e.RowIndex].Cells[14].Value.ToString();
                    string resim = advancedDataGridView1.Rows[e.RowIndex].Cells[15].Value.ToString();
                    string HatayıBulanBolum = advancedDataGridView1.Rows[e.RowIndex].Cells[16].Value.ToString();
                    string KokNeden = advancedDataGridView1.Rows[e.RowIndex].Cells[17].Value.ToString();
                    string Aksiyon = advancedDataGridView1.Rows[e.RowIndex].Cells[18].Value.ToString();
                    string Sonuc = advancedDataGridView1.Rows[e.RowIndex].Cells[19].Value.ToString();

                    UOD.Resim1 = resim; // Form2'deki TextBox'a değeri aktar
                    UOD.UrunKodu1 = UrunKodu;
                    UOD.UrunAdi1 = UrunAdi;
                    UOD.SipNo1 = SipNo;
                    UOD.HataMik1 = HataMik;
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
                    UOD.ButtonGuncelle = true;
                    UOD.Show();
                    this.Hide();
                }
            }
        }

        private void UygunOlmayanListe_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Form kapatılmak istendiğinde
            if (e.CloseReason == CloseReason.UserClosing)
            {
                // Programı kapat
                Application.Exit();
            }
        }
        Resim RS;
        private void advancedDataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 14)
            {
                if (RS == null || RS.IsDisposed)
                {
                    RS = new Resim();
                    string resim = advancedDataGridView1.Rows[e.RowIndex].Cells[14].Value.ToString();
                    RS.picturebox132 = resim;
                    RS.Show();

                }
            }
        }

        private void advancedDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 20) // 19. sütundaki düğme tıklandı mı kontrol ediliyor
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
                    eskiHataliUrun.Durum = "False";

                    // Değişiklikleri veritabanına kaydedin
                    dbContext.SaveChanges();
                    MessageBox.Show("Güncellendi.");
                    Uygunolmayanvol1();
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
    }
}
