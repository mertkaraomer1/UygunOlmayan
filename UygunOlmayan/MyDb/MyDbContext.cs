using UygunOlmayan.Tables;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UygunOlmayan.MyDb
{
    public class MyDbContext:DbContext
    {

        public DbSet<HataliUrun> hataliUruns { get; set; }
        public DbSet<HataGrupları> hataGruplars { get; set; }
        public DbSet<UrunTıpı> urunTıpıs {  get; set; }
        public DbSet<Sorun_Bildirim> sorun_Bildirims { get; set; }
        public DbSet<AksiyonAlacakBölüm> aksiyonAlacakBölüms { get; set; }



        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // Burada veritabanı bağlantı bilgilerini tanımlayın.
            // Örnek olarak SQL Server kullanalım:
            string connectionString = "Data Source=192.168.2.250;Initial Catalog=Muh_Plan_Prog1;Integrated Security=True;Connect Timeout=10;Encrypt=False;Trust Server Certificate=False;Application Intent=ReadWrite;Multi Subnet Failover=False";
            optionsBuilder.UseSqlServer(connectionString);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            modelBuilder.Entity<HataliUrun>().ToTable("HataliUrun").HasKey(x => x.UrunId);
            modelBuilder.Entity<HataGrupları>().ToTable("HataGrupları").HasKey(x=>x.HataId);
            modelBuilder.Entity<UrunTıpı>().ToTable("UrunTıpı").HasKey(x => x.Id);
            modelBuilder.Entity<Sorun_Bildirim>().ToTable("Sorun_Bildirim").HasKey(x => x.Sorun_ID);
            modelBuilder.Entity<AksiyonAlacakBölüm>().ToTable("AksiyonAlacakBölüm").HasKey(x => x.UygunOlmayanId);

        }
    }

}
