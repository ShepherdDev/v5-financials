using Rock.Plugin;

namespace com.shepherdchurch.v5_financials.Migrations
{
    [MigrationNumber( 3, "1.5.0" )]
    public class ConfigureBlock : Migration
    {
        static string ExportToGLBlockTypeGuid = "3f161c8b-2d36-4be7-a714-7fd8f6f6f1ec";
        static string ExportToGLBlockGuid = "1b803337-4c16-40a5-a932-b3b2b7d7b870";
        static string BatchDetailPageGuid = "606BDA31-A8FE-473A-B3F8-A00ECF7E06EC";

        /// <summary>
        /// The commands to run to migrate plugin to the specific version
        /// </summary>
        public override void Up()
        {
            RockMigrationHelper.UpdateBlockType( "Export to GL", "Export the current batch to the GL system.", "~/Plugins/com_shepherdchurch/v5-financials/ExportToGL.ascx", "com_shepherdchurch > V5 Financials", ExportToGLBlockTypeGuid );
            RockMigrationHelper.AddBlock( BatchDetailPageGuid, "", ExportToGLBlockTypeGuid, "Export To GL", "Main", string.Empty, string.Empty, 100, ExportToGLBlockGuid );
        }

        /// <summary>
        /// The commands to undo a migration from a specific version
        /// </summary>
        public override void Down()
        {
            RockMigrationHelper.DeleteBlock( ExportToGLBlockGuid );
            RockMigrationHelper.DeleteBlockType( ExportToGLBlockTypeGuid );
        }
    }
}

