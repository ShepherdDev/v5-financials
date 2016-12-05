using Rock.Plugin;

namespace com.shepherdchurch.v5_financials.Migrations
{
    [MigrationNumber( 2, "1.5.0" )]
    public class ExportedAttribute : Migration
    {
        static string GLExportedGuid = "3324e5ee-3c34-4786-a74c-ad863147560a";

        /// <summary>
        /// The commands to run to migrate plugin to the specific version
        /// </summary>
        public override void Up()
        {
            RockMigrationHelper.AddEntityAttribute( "Rock.Model.FinancialBatch", "9C204CD0-1233-41C5-818A-C5DA439445AA", "", "",
                "GL Exported", string.Empty, "True if this batch has been exported to GL.", 100, "false", GLExportedGuid );
        }

        /// <summary>
        /// The commands to undo a migration from a specific version
        /// </summary>
        public override void Down()
        {
            RockMigrationHelper.DeleteAttribute( GLExportedGuid );
        }
    }
}

