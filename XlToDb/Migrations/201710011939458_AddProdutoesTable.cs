namespace XlToDb.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddProdutoesTable : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Produtoes",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Apelido = c.String(maxLength: 10),
                        Descricao = c.String(maxLength: 100),
                        UnidadeId = c.Int(nullable: false),
                        TipoId = c.Int(nullable: false),
                        ClasseCustoId = c.Int(nullable: false),
                        FamiliaId = c.Int(nullable: false),
                        LinhaId = c.Int(nullable: false),
                        GrupoRateioId = c.Int(nullable: false),
                        CategoriaId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.Categorias", t => t.CategoriaId, cascadeDelete: true)
                .ForeignKey("dbo.ClasseCustoes", t => t.ClasseCustoId, cascadeDelete: true)
                .ForeignKey("dbo.Familias", t => t.FamiliaId, cascadeDelete: true)
                .ForeignKey("dbo.GrupoRateios", t => t.GrupoRateioId, cascadeDelete: true)
                .ForeignKey("dbo.Linhas", t => t.LinhaId, cascadeDelete: true)
                .ForeignKey("dbo.Tipoes", t => t.TipoId, cascadeDelete: true)
                .ForeignKey("dbo.Unidades", t => t.UnidadeId, cascadeDelete: true)
                .Index(t => t.UnidadeId)
                .Index(t => t.TipoId)
                .Index(t => t.ClasseCustoId)
                .Index(t => t.FamiliaId)
                .Index(t => t.LinhaId)
                .Index(t => t.GrupoRateioId)
                .Index(t => t.CategoriaId);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.Produtoes", "UnidadeId", "dbo.Unidades");
            DropForeignKey("dbo.Produtoes", "TipoId", "dbo.Tipoes");
            DropForeignKey("dbo.Produtoes", "LinhaId", "dbo.Linhas");
            DropForeignKey("dbo.Produtoes", "GrupoRateioId", "dbo.GrupoRateios");
            DropForeignKey("dbo.Produtoes", "FamiliaId", "dbo.Familias");
            DropForeignKey("dbo.Produtoes", "ClasseCustoId", "dbo.ClasseCustoes");
            DropForeignKey("dbo.Produtoes", "CategoriaId", "dbo.Categorias");
            DropIndex("dbo.Produtoes", new[] { "CategoriaId" });
            DropIndex("dbo.Produtoes", new[] { "GrupoRateioId" });
            DropIndex("dbo.Produtoes", new[] { "LinhaId" });
            DropIndex("dbo.Produtoes", new[] { "FamiliaId" });
            DropIndex("dbo.Produtoes", new[] { "ClasseCustoId" });
            DropIndex("dbo.Produtoes", new[] { "TipoId" });
            DropIndex("dbo.Produtoes", new[] { "UnidadeId" });
            DropTable("dbo.Produtoes");
        }
    }
}
