namespace ExcelTest.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InitialCreate : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Statistics",
                c => new
                    {
                        StatisticId = c.Int(nullable: false, identity: true),
                        StartDate = c.DateTime(nullable: false),
                        EndDate = c.DateTime(nullable: false),
                        CreationDate = c.DateTime(nullable: false),
                        Creator = c.String(),
                        Car_Id = c.Int(),
                    })
                .PrimaryKey(t => t.StatisticId)
                .ForeignKey("dbo.Cars", t => t.Car_Id)
                .Index(t => t.Car_Id);
            
            CreateTable(
                "dbo.Cars",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Description = c.String(),
                        Radio = c.String(),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.Statistics", "Car_Id", "dbo.Cars");
            DropIndex("dbo.Statistics", new[] { "Car_Id" });
            DropTable("dbo.Cars");
            DropTable("dbo.Statistics");
        }
    }
}
