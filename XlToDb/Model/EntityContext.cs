using System.Data.Entity;

namespace XlToDb.Model
{
    class EntityContext : DbContext
    {
        public EntityContext() : base("name=SqlServer")
        {
        }

        public DbSet<Categoria> Categorias { get; set; }
        public DbSet<ClasseCusto> ClassesCusto { get; set; }
        public DbSet<Familia> Familias { get; set; }
        public DbSet<GrupoRateio> GruposRateio { get; set; }
        public DbSet<Linha> Linhas { get; set; }
        public DbSet<LogData> LogData { get; set; }
        public DbSet<Produto> Produtos { get; set; }
        public DbSet<Tipo> Tipos { get; set; }
        public DbSet<Unidade> Unidades { get; set; }
        public DbSet<Estrutura> Estruturas { get; set; }
        public DbSet<Setor> Setores { get; set; }
        public DbSet<Operacao> Operacoes { get; set; }
        public DbSet<Area> Areas { get; set; }
        public DbSet<CustoFolha> CustoFolhas { get; set; }
        public DbSet<MedidaFita> MedidaFitas { get; set; }
        public DbSet<QtdEmbalagem> QtdEmbalagems { get; set; }
        public DbSet<Finalidade> Finalidades { get; set; }
        public DbSet<Insumo> Insumos { get; set; }
    }
}
