using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UygunOlmayan.Tables
{
    public class HataliUrun
    {
        public int UrunId { get; set; }
        public string UrunKodu { get; set; }
        public string? UrunAdi { get; set; }
        public string SiparisNo { get; set; }
        public int? HatalıMiktar { get; set; }
        public string? Adet { get; set; }
        public DateTime Tarih { get; set; }
        public int? KayıpZaman { get; set; }
        public string? ZamanCinsi { get; set; }
        public string? HataTipi { get; set; }
        public string? Aciklama { get; set; }
        public string? Ozet { get; set; }
        public string? HataBolumu { get; set; }
        public string? RaporuHazirlayan { get; set; }
        public string? HatayıBulanBirim { get; set; }
        public string? Resim { get; set; }
        public string? KokNeden { get; set; }
        public string? Aksiyon { get; set; }
        public string? Sonuc { get; set; }
        public  string? Durum { get; set; }
        public string? Tedarikci { get; set; }  

    }
}
