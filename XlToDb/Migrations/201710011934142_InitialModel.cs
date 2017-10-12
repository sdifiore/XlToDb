namespace XlToDb.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InitialModel : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Categorias",
                c => new
                    {
                        CategoriaId = c.Int(nullable: false, identity: true),
                        Apelido = c.String(maxLength: 2),
                        Descricao = c.String(maxLength: 50),
                    })
                .PrimaryKey(t => t.CategoriaId);
            
            CreateTable(
                "dbo.ClasseCustoes",
                c => new
                    {
                        ClasseCustoId = c.Int(nullable: false, identity: true),
                        Apelido = c.String(maxLength: 2),
                        Descricao = c.String(maxLength: 50),
                    })
                .PrimaryKey(t => t.ClasseCustoId);
            
            CreateTable(
                "dbo.Familias",
                c => new
                    {
                        FamiliaId = c.Int(nullable: false, identity: true),
                        Apelido = c.String(maxLength: 3),
                        Descricao = c.String(maxLength: 50),
                    })
                .PrimaryKey(t => t.FamiliaId);
            
            CreateTable(
                "dbo.GrupoRateios",
                c => new
                    {
                        GrupoRateioId = c.Int(nullable: false, identity: true),
                        Descricao = c.String(maxLength: 50),
                    })
                .PrimaryKey(t => t.GrupoRateioId);
            
            CreateTable(
                "dbo.Linhas",
                c => new
                    {
                        LinhaId = c.Int(nullable: false, identity: true),
                        Apelido = c.String(maxLength: 4),
                        Descricao = c.String(maxLength: 50),
                    })
                .PrimaryKey(t => t.LinhaId);
            
            CreateTable(
                "dbo.Tipoes",
                c => new
                    {
                        TipoId = c.Int(nullable: false, identity: true),
                        Apelido = c.String(maxLength: 1),
                        Descricao = c.String(maxLength: 50),
                    })
                .PrimaryKey(t => t.TipoId);
            
            CreateTable(
                "dbo.Unidades",
                c => new
                    {
                        UnidadeId = c.Int(nullable: false, identity: true),
                        Apelido = c.String(maxLength: 2),
                        Descricao = c.String(maxLength: 16),
                    })
                .PrimaryKey(t => t.UnidadeId);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.Unidades");
            DropTable("dbo.Tipoes");
            DropTable("dbo.Linhas");
            DropTable("dbo.GrupoRateios");
            DropTable("dbo.Familias");
            DropTable("dbo.ClasseCustoes");
            DropTable("dbo.Categorias");
        }
    }
}
