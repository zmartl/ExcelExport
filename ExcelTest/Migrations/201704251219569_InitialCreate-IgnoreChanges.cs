namespace ExcelTest.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InitialCreateIgnoreChanges : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Plannings",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        StartTime = c.DateTime(nullable: false),
                        EndTime = c.DateTime(nullable: false),
                        Car_Id = c.Int(),
                        State_Id = c.Int(),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.Cars", t => t.Car_Id)
                .ForeignKey("dbo.States", t => t.State_Id)
                .Index(t => t.Car_Id)
                .Index(t => t.State_Id);
            
            CreateTable(
                "dbo.States",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.Plannings", "State_Id", "dbo.States");
            DropForeignKey("dbo.Plannings", "Car_Id", "dbo.Cars");
            DropIndex("dbo.Plannings", new[] { "State_Id" });
            DropIndex("dbo.Plannings", new[] { "Car_Id" });
            DropTable("dbo.States");
            DropTable("dbo.Plannings");
        }
    }
}
