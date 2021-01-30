namespace AICO_CL.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class newdatabase : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Accountings",
                c => new
                    {
                        ID = c.Int(nullable: false, identity: true),
                        EmployeID = c.Int(),
                        ComputerID = c.Int(),
                        DeviceID = c.Int(),
                    })
                .PrimaryKey(t => t.ID)
                .ForeignKey("dbo.Computers", t => t.ComputerID)
                .ForeignKey("dbo.Devices", t => t.DeviceID)
                .ForeignKey("dbo.Employes", t => t.EmployeID)
                .Index(t => t.EmployeID)
                .Index(t => t.ComputerID)
                .Index(t => t.DeviceID);
            
            CreateTable(
                "dbo.Computers",
                c => new
                    {
                        ID = c.Int(nullable: false, identity: true),
                        UserNamePC = c.String(),
                        NamePC = c.String(),
                        OSVersion = c.String(),
                        BitOperating = c.String(),
                        Motherboard = c.String(),
                        CPUpc = c.String(),
                        RAMpc = c.String(),
                        HDDpc = c.String(),
                        Video = c.String(),
                    })
                .PrimaryKey(t => t.ID);
            
            CreateTable(
                "dbo.Devices",
                c => new
                    {
                        ID = c.Int(nullable: false, identity: true),
                        Model = c.String(),
                        Description_1 = c.String(),
                        Description_2 = c.String(),
                        Description_3 = c.String(),
                        Description_4 = c.String(),
                        Description_5 = c.String(),
                        Device_ENUM_ID = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.ID)
                .ForeignKey("dbo.Device_ENUM", t => t.Device_ENUM_ID, cascadeDelete: true)
                .Index(t => t.Device_ENUM_ID);
            
            CreateTable(
                "dbo.Device_ENUM",
                c => new
                    {
                        ID = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.ID);
            
            CreateTable(
                "dbo.Employes",
                c => new
                    {
                        ID = c.Int(nullable: false, identity: true),
                        Work = c.String(),
                        Name = c.String(),
                        Password = c.String(),
                        Phone = c.String(),
                        DepartmentID = c.Int(),
                    })
                .PrimaryKey(t => t.ID)
                .ForeignKey("dbo.Departments", t => t.DepartmentID)
                .Index(t => t.DepartmentID);
            
            CreateTable(
                "dbo.Departments",
                c => new
                    {
                        ID = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.ID);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.Accountings", "EmployeID", "dbo.Employes");
            DropForeignKey("dbo.Employes", "DepartmentID", "dbo.Departments");
            DropForeignKey("dbo.Accountings", "DeviceID", "dbo.Devices");
            DropForeignKey("dbo.Devices", "Device_ENUM_ID", "dbo.Device_ENUM");
            DropForeignKey("dbo.Accountings", "ComputerID", "dbo.Computers");
            DropIndex("dbo.Employes", new[] { "DepartmentID" });
            DropIndex("dbo.Devices", new[] { "Device_ENUM_ID" });
            DropIndex("dbo.Accountings", new[] { "DeviceID" });
            DropIndex("dbo.Accountings", new[] { "ComputerID" });
            DropIndex("dbo.Accountings", new[] { "EmployeID" });
            DropTable("dbo.Departments");
            DropTable("dbo.Employes");
            DropTable("dbo.Device_ENUM");
            DropTable("dbo.Devices");
            DropTable("dbo.Computers");
            DropTable("dbo.Accountings");
        }
    }
}
