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
            table.Columns.Add("KayıpGun");  // kayıp gün kolonu da şart!

            var hataliUrunList = dbContext.hataliUruns
                .Where(x => x.HataBolumu == HatanınOlustuguBolum&& x.urunimza != new Guid())
                .Select(x => new
                {
                    x.HataTipi,
                    x.HataBolumu,
                    x.SiparisNo,
                    x.UrunKodu,
                    x.KokNeden,
                    x.KokNedenAksiyon,
                    kayıpGun =Math.Round( (x.KapanısTarihi - x.Tarih).TotalDays,0 ) // sadece gün sayısını verir
                })
                .ToList();
            int sayac = 1;  
            foreach (var urun in hataliUrunList)
            {
                table.Rows.Add(sayac++,urun.UrunKodu, urun.SiparisNo,  urun.HataTipi,urun.HataBolumu, urun.KokNeden,urun.KokNedenAksiyon,urun.kayıpGun);
            }
            advancedDataGridView1.DataSource = table;
        }
    }
}
