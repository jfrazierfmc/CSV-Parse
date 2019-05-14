namespace Vidly.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class SeedUsers : DbMigration
    {
        public override void Up()
        {
            Sql(@"INSERT INTO [dbo].[AspNetUsers] ([Id], [Email], [EmailConfirmed], [PasswordHash], [SecurityStamp], [PhoneNumber], [PhoneNumberConfirmed], [TwoFactorEnabled], [LockoutEndDateUtc], [LockoutEnabled], [AccessFailedCount], [UserName]) VALUES (N'8592f255-1120-47cf-94b6-8c73b8956600', N'admin@vidly.com', 0, N'AIrZ95UHkzNY1xb6jffhFP9l/pMa2c2L+XJDzmXbIVb3lcW5QBjXg5SlZEMoXU1nNA==', N'1a507721-c3fb-4c1d-b870-faeb97baf0fc', NULL, 0, 0, NULL, 1, 0, N'admin@vidly.com')
INSERT INTO [dbo].[AspNetUsers] ([Id], [Email], [EmailConfirmed], [PasswordHash], [SecurityStamp], [PhoneNumber], [PhoneNumberConfirmed], [TwoFactorEnabled], [LockoutEndDateUtc], [LockoutEnabled], [AccessFailedCount], [UserName]) VALUES (N'9c9dfef3-ab0b-4367-a6a1-94f42871d317', N'guest@vidly.com', 0, N'AIR1CeEIj+GtqyoCSWY7WykHJx7kFwVXYLYzTfzcfCGrR0DtGHU5HMMnRgzq0B1AOg==', N'ab44b22c-65e1-40e1-8595-00cfe19d35f1', NULL, 0, 0, NULL, 1, 0, N'guest@vidly.com')

INSERT INTO [dbo].[AspNetRoles] ([Id], [Name]) VALUES (N'583a534d-00c3-4d1a-a433-9b1ebe32c344', N'CanManageMovies')

INSERT INTO [dbo].[AspNetUserRoles] ([UserId], [RoleId]) VALUES (N'8592f255-1120-47cf-94b6-8c73b8956600', N'583a534d-00c3-4d1a-a433-9b1ebe32c344')
");

        }

    public override void Down()
        {
        }
    }
}
