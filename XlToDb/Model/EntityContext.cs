using System.Data.Entity;

namespace XlToDb.Model
{
    class EntityContext : DbContext
    {
        public EntityContext() : base("name=SqlServer")
        {
        }

        public DbSet<Categoria> Categorias { get; set; }
        public DbSet<Base> Bases { get; set; }
        public DbSet<ClasseCusto> ClassesCusto { get; set; }
        public DbSet<CondicaoPreco> CondicoesPrecos { get; set; }
        public DbSet<Cotacao> Cotacoes { get; set; }
        public DbSet<CuboEstoque> CubosEstoque { get; set; }
        public DbSet<DfxProdRev> DfxProdRevs { get; set; }
        public DbSet<Embal> Embals { get; set; }
        public DbSet<Embalagem> Embalagens { get; set; }
        public DbSet<Familia> Familias { get; set; }
        public DbSet<Fabricante> Fabricantes { get; set; }
        public DbSet<GrupoRateio> GruposRateio { get; set; }
        public DbSet<Linha> Linhas { get; set; }
        public DbSet<LogData> LogData { get; set; }
        public DbSet<Produto> Produtos { get; set; }
        public DbSet<Tipo> Tipos { get; set; }
        public DbSet<Unidade> Unidades { get; set; }
        public DbSet<Estrutura> Estruturas { get; set; }
        public DbSet<Serie> Series { get; set; }
        public DbSet<Setor> Setores { get; set; }
        public DbSet<Operacao> Operacoes { get; set; }
        public DbSet<Area> Areas { get; set; }
        public DbSet<Carga> Cargas { get; set; }
        public DbSet<CustoCargoDireto> CustoCargoDiretos { get; set; }
        public DbSet<CustoFolha> CustoFolhas { get; set; }
        public DbSet<MedidaFita> MedidaFitas { get; set; }
        public DbSet<QtdEmbalagem> QtdEmbalagems { get; set; }
        public DbSet<Finalidade> Finalidades { get; set; }
        public DbSet<Insumo> Insumos { get; set; }
        public DbSet<InsumoXtd> InsumoXtds { get; set; }
        public DbSet<Dominio> Dominios { get; set; }
        public DbSet<DespesaFixa> DespesasFixas { get; set; }
        public DbSet<ParteProduto> ParteProdutos { get; set; }
        public DbSet<PlanejCompra> PlanejCompras { get; set; }
        public DbSet<PlanejMod> PlanejMods { get; set; }
        public DbSet<PrecoNacional> PrecosNacionais { get; set; }
        public DbSet<PrecoExportacao> PrecosExportacao { get; set; }
        public DbSet<ProcTubo> ProcTubos { get; set; }
        public DbSet<Ajuste> Ajustes { get; set; }
        public DbSet<TipoAlteracao> TiposAlteracao { get; set; }
        public DbSet<EncapTubo> EncapTubos { get; set; }
        public DbSet<Graxa> Graxas { get; set; }
        public DbSet<Resina> Resinas { get; set; }
        public DbSet<ResinaBase> ResinasBase { get; set; }
        public DbSet<ResinaPtfe> ResinasPtfe { get; set; }
        public DbSet<PreForma> PreFormas { get; set; }
        public DbSet<Extrusora> Extrusoras { get; set; }
        public DbSet<PrensaPreForma> PrensasPreForma { get; set; }
        public DbSet<PadraoFixo> PadroesFixos { get; set; }
        public DbSet<AjusteProduto> AjusteProdutos { get; set; }



    }
}
