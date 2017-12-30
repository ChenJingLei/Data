namespace Data.Service
{
    public class ApplicationDbContext : DbContext
    {
        public DbSet<Export> Exports { get; set; }
        public DbSet<Info> Infos { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(@"Server=(localdb)\mssqllocaldb;Database=MyDatabase;Trusted_Connection=True;");
        }
    }
}