using Rock.Plugin;

namespace com.shepherdchurch.v5_financials.Migrations
{
    [MigrationNumber( 1, "1.5.0" )]
    public class Attributes : Migration
    {
        static string GLCompanyGuid = "290fed3a-80e2-49ca-b8ce-4c4210699ea4";
        static string GLFundGuid = "79743425-a0c5-43ce-ab9b-ee63c0341783";
        static string GLBankAccountGuid = "1ce18f2d-6d79-41cf-ab65-a6a228ea1a7e";
        static string GLRevenueDepartmentGuid = "2ab2cb87-6b7a-4b54-b592-f9a61f4a4aee";
        static string GLRevenueAccountGuid = "7810b253-a53b-4f2a-932c-dff25971f1a8";

        /// <summary>
        /// The commands to run to migrate plugin to the specific version
        /// </summary>
        public override void Up()
        {
            RockMigrationHelper.AddEntityAttribute( "Rock.Model.FinancialAccount", "9C204CD0-1233-41C5-818A-C5DA439445AA", "", "",
                "GL Company", string.Empty, "3 digit company code that this account is associated with.", 100, "", GLCompanyGuid );

            RockMigrationHelper.AddEntityAttribute( "Rock.Model.FinancialAccount", "9C204CD0-1233-41C5-818A-C5DA439445AA", "", "",
                "GL Fund", string.Empty, "3 digit code for the fund that this account is associated with.", 101, "", GLFundGuid );

            RockMigrationHelper.AddEntityAttribute( "Rock.Model.FinancialAccount", "9C204CD0-1233-41C5-818A-C5DA439445AA", "", "",
                "GL Bank Account", string.Empty, "9 digit account number that identifies the bank account in the GL system that this account is associated with. This is where the money is deposited into.", 102, "", GLBankAccountGuid );

            RockMigrationHelper.AddEntityAttribute( "Rock.Model.FinancialAccount", "9C204CD0-1233-41C5-818A-C5DA439445AA", "", "",
                "GL Revenue Department", string.Empty, "3 digit department code for this account.", 103, "", GLRevenueDepartmentGuid );

            RockMigrationHelper.AddEntityAttribute( "Rock.Model.FinancialAccount", "9C204CD0-1233-41C5-818A-C5DA439445AA", "", "",
                "GL Revenue Account", string.Empty, "9 digit account number that identifies the revenue account in the GL system that this account is associated with. This is where the money is withdrawn from.", 104, "", GLRevenueAccountGuid );
        }

        /// <summary>
        /// The commands to undo a migration from a specific version
        /// </summary>
        public override void Down()
        {
            RockMigrationHelper.DeleteAttribute( GLRevenueAccountGuid );
            RockMigrationHelper.DeleteAttribute( GLRevenueDepartmentGuid );
            RockMigrationHelper.DeleteAttribute( GLBankAccountGuid );
            RockMigrationHelper.DeleteAttribute( GLFundGuid );
            RockMigrationHelper.DeleteAttribute( GLCompanyGuid );
        }
    }
}

