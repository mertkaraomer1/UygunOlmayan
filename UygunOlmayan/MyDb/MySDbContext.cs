using Microsoft.EntityFrameworkCore;
using UygunOlmayan.Tables;


namespace UygunOlmayan.MyDb
{
    public class MySDbContext : DbContext
    {
        public DbSet<STOKLAR> STOKLAR { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // Burada veritabanı bağlantı bilgilerini tanımlayın.
            // Örnek olarak SQL Server kullanalım:
            string connectionString = "Data Source=SRV-MIKRO;Initial Catalog=MikroDB_V16_ICM;Integrated Security=True;Connect Timeout=10;Encrypt=False;Trust Server Certificate=False;Application Intent=ReadWrite;Multi Subnet Failover=False";
            optionsBuilder.UseSqlServer(connectionString);
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            modelBuilder.Entity<STOKLAR>().ToTable("STOKLAR").HasKey(x => x.sto_Guid);

        }

    }
}
