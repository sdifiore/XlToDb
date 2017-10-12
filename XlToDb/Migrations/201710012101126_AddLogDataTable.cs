namespace XlToDb.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddLogDataTable : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.LogData",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Timestamp = c.DateTime(nullable: false),
                        Reason = c.String(nullable: false, maxLength: 5),
                        Message = c.String(nullable: false, maxLength: 256),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.LogData");
        }
    }
}
