// Plugin namespaces
using System;
using System.ServiceModel;

// Microsoft Dynamics CRM namespace(s)
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Collections.Generic;

// Sample plugin structure adapted but all code modified
namespace Microsoft.Crm.Sdk.Samples
{
    public class goMDS : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {
            // Extract the tracing service for use in debugging sandboxed plug-ins
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters
                Entity entity = (Entity)context.InputParameters["Target"];

                // Verify that the target entity represents the Palm Go MDS entity
                // If not, this plug-in was not registered correctly.
                // if (entity.LogicalName != "new_palmgomds")
                //    return;

                try
                {
                    // Fixes:
                    // Remove ampersand from comments
                    // Compare values to Lowercase

                    // Global variables
                    string varDescription = ""; // Description field on form
                    bool varCreateExtract = false; // Create Extract field on form
                    int varYear = 0; // Year from form
                    int varQuarter = 0; // Quarter from form
                    int varQuarterFull = 0; // Quarter expressed as optionset id
                    Guid varMdsID = new Guid(); // GUID for palm go mds record
                    StringBuilder sbHeaderList = new StringBuilder(); // String builder for extract header
                    StringBuilder sbClientList = new StringBuilder(); // String builder for extract client data
                    StringBuilder sbErrorList = new StringBuilder(); // String builder for error list
                    string varFileName = ""; // Extract file name
                    string varFileName2 = ""; // Error log file name
                    DateTime varStartDate = new DateTime(); // Start date of extract
                    DateTime varEndDate = new DateTime(); // End date of extract
                    int varCheckInt = 0; // Used to see if data is valid integer
                    double varCheckDouble = 0;  // Used to see if data is valid double
                    DateTime varCheckDate = new DateTime(); // Used to see if data is valid date
                    EntityReference getEntity; // Object to get entity details

                    string varTest = ""; // Used for debug

                    // Only do this if the entity is the Palm Go MDS entity
                    if (entity.LogicalName == "new_palmgomds")
                    {
                        // Login to system as user
                        if (Equals(serviceProvider, null))
                        {
                            throw new ArgumentNullException("serviceProvider");
                        }

                        // Objects required for the plugin and to get the organisation / user details
                        IOrganizationService _service;
                        IPluginExecutionContext PluginExecutionContext = serviceProvider.GetService(typeof(IPluginExecutionContext)) as IPluginExecutionContext;

                        // Obtain the organisation service factory from the service provider
                        var factory = serviceProvider.GetService(typeof(IOrganizationServiceFactory)) as IOrganizationServiceFactory;
                        if (factory == null)
                        {
                            throw new InvalidPluginExecutionException("Unable to get OrganizationServiceFactory");
                        }

                        // Use the factory to generate the Organisation Service
                        if (PluginExecutionContext != null && (!Equals(PluginExecutionContext.UserId) && PluginExecutionContext.UserId != Guid.Empty))
                        {
                            _service = factory.CreateOrganizationService(PluginExecutionContext.UserId);
                        }
                        else
                        {
                            // Impersonate as system
                            _service = factory.CreateOrganizationService(null);
                        }

                        // Get info from current Palm Go MDS record
                        varDescription = entity.GetAttributeValue<string>("new_description");
                        varCreateExtract = entity.GetAttributeValue<bool>("new_createextract");
                        varYear = entity.GetAttributeValue<int>("new_year");
                        varQuarter = entity.GetAttributeValue<OptionSetValue>("new_quarter").Value;
                        varMdsID = entity.Id;
                        // Assign Quarter Full to the Option Set value, then subtract 100000000 to get the normal quarter value
                        varQuarterFull = varQuarter;
                        varQuarter = varQuarter - 100000000;

                        // Create the file names
                        varFileName = "HACC__03580" + varYear + varQuarter + "0101.csv";
                        varFileName2 = "Errors for HACC__03580" + varYear + varQuarter + "0101.txt";

                        //Get the start and end date of period
                        if (varQuarter == 1)
                        {
                            varStartDate = Convert.ToDateTime("1-Jan-" + varYear);
                            varEndDate = Convert.ToDateTime("31-Mar-" + varYear);
                        }
                        else if (varQuarter == 2)
                        {
                            varStartDate = Convert.ToDateTime("1-Apr-" + varYear);
                            varEndDate = Convert.ToDateTime("30-Jun-" + varYear);
                        }
                        else if (varQuarter == 3)
                        {
                            varStartDate = Convert.ToDateTime("1-Jul-" + varYear);
                            varEndDate = Convert.ToDateTime("30-Sep-" + varYear);
                        }
                        else
                        {
                            varStartDate = Convert.ToDateTime("1-Oct-" + varYear);
                            varEndDate = Convert.ToDateTime("31-Dec-" + varYear);
                        }

                        // Important: The plugin uses American dates but returns formatted Australian dates
                        // Any dates created in the plugin will be American
                        varEndDate = varEndDate.AddHours(23);

                        // varTest += varStartDate + " " + varEndDate + "\r\n";

                        // Fetch statements for database
                        // Get the required fields from the client table (and associated entities)
                        // Any clients that have an MDS record for the current quarter
                        string dbMDSList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclient'>
                                <attribute name='new_address' />
                                <attribute name='new_country' />
                                <attribute name='new_depchild' />
                                <attribute name='new_dob' />
                                <attribute name='new_dobest' />
                                <attribute name='new_firstname' />
                                <attribute name='new_gender' />
                                <attribute name='new_sex' />
                                <attribute name='new_interpret' />
                                <attribute name='new_indigenous' />
                                <attribute name='new_language' />
                                <attribute name='new_mdsslk' />
                                <attribute name='new_surname' />
                                <attribute name='new_palmclientid' />
                                <link-entity name='new_palmclientmds' to='new_palmclientid' from='new_client' link-type='inner'>
                                    <attribute name='new_year' />
                                    <attribute name='new_quarter' />
                                </link-entity>
                                <link-entity name='new_palmddlcountry' to='new_country' from='new_palmddlcountryid' link-type='outer'>
                                    <attribute name='new_code' />
                                </link-entity>
                                <link-entity name='new_palmddllanguage' to='new_language' from='new_palmddllanguageid' link-type='outer'>
                                    <attribute name='new_code' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientmds' attribute='new_year' operator='eq' value='" + varYear + @"' />
                                    <condition entityname='new_palmclientmds' attribute='new_quarter' operator='eq' value='" + varQuarterFull + @"' />
                                </filter>
                                <order attribute='new_address' />
                              </entity>
                            </fetch> ";

                        // Get the required fields from the SU Drop Down list entity
                        string dbDropList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                  <entity name='new_palmsudropdown'>
                                    <attribute name='new_type' />
                                    <attribute name='new_description' />
                                    <attribute name='new_mds' />
                                    <order attribute='new_description' />
                                  </entity>
                                </fetch> ";

                        // Get the required fields from the support period (and associated entities)
                        // Any support periods active during the period checked as MDS, where the client has an MDS record for the quarter
                        string dbSupportList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                    <entity name='new_palmclientsupport'>
                                        <attribute name='modifiedon' />
                                        <attribute name='new_careravail' />
                                        <attribute name='new_carercountry' />
                                        <attribute name='new_carerdob' />
                                        <attribute name='new_carerdobflag' />
                                        <attribute name='new_carerfirstname' />
                                        <attribute name='new_carerindigenous' />
                                        <attribute name='new_carerlanguage' />
                                        <attribute name='new_carerlocality' />
                                        <attribute name='new_carermore' />
                                        <attribute name='new_carerresidence' />
                                        <attribute name='new_carerrship' />
                                        <attribute name='new_carersex' />
                                        <attribute name='new_carerslk' />
                                        <attribute name='new_carersurname' />
                                        <attribute name='new_cessation' />
                                        <attribute name='new_client' />
                                        <attribute name='new_ccpdistype' />
                                        <attribute name='new_dva' />
                                        <attribute name='new_enddate' />
                                        <attribute name='new_incomepres' />
                                        <attribute name='new_livingarrangepres' />
                                        <attribute name='new_locality' />
                                        <attribute name='new_residentialpres' />
                                        <attribute name='new_sourceref' />
                                        <attribute name='new_startdate' />
                                        <link-entity name='new_palmclientmds' to='new_client' from='new_client' link-type='inner'>
                                            <attribute name='new_year' />
                                            <attribute name='new_quarter' />
                                        </link-entity>
                                        <link-entity name='new_palmddlcountry' to='new_carercountry' from='new_palmddlcountryid' link-type='outer'>
                                            <attribute name='new_code' />
                                        </link-entity>
                                        <link-entity name='new_palmddllanguage' to='new_carerlanguage' from='new_palmddllanguageid' link-type='outer'>
                                            <attribute name='new_code' />
                                        </link-entity>
                                        <link-entity name='new_palmddllocality' to='new_locality' from='new_palmddllocalityid' link-type='outer'>
                                            <attribute name='new_postcode' />
                                            <attribute name='new_state' />
                                        </link-entity>
                                        <link-entity name='new_palmddllocality' to='new_carerlocality' from='new_palmddllocalityid' link-type='outer'>
                                            <attribute name='new_postcode' />
                                            <attribute name='new_state' />
                                        </link-entity>
                                        <filter type='and'>
                                            <condition entityname='new_palmclientmds' attribute='new_year' operator='eq' value='" + varYear + @"' />
                                            <condition entityname='new_palmclientmds' attribute='new_quarter' operator='eq' value='" + varQuarterFull + @"' />
                                            <condition entityname='new_palmclientsupport' attribute='new_mds' operator='eq' value='True' />
                                            <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                            <filter type='or'>
                                                <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                                <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                            </filter >
                                        </filter>
                                        <order attribute='new_client' />
                                        <order attribute='modifiedon' descending='true' />
                                        <order attribute='new_startdate' descending='true' />
                                    </entity>
                                </fetch> ";

                        // Get the required fields from the Activities entity
                        // Any activities marked as MDS against a support period marked as MDS, where the activity occurs in the quarter and the support period is active during the quarter
                        string dbActivityList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                    <entity name='new_palmclientactivities'>
                                        <attribute name='new_supportperiod' />
                                        <attribute name='new_amount' />
                                        <attribute name='new_activitymds' />
                                        <attribute name='new_entrydate' />
                                        <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                                <attribute name='new_client' />
                                        </link-entity>
                                        <filter type='and'>
                                            <condition entityname='new_palmclientactivities' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                            <condition entityname='new_palmclientactivities' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                            <condition entityname='new_palmclientsupport' attribute='new_mds' operator='eq' value='True' />
                                            <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                            <filter type='or'>
                                                <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                                <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                            </filter >
                                        </filter>
                                    </entity>
                                </fetch> ";

                        // Get the required fields from the Financials entity
                        // Any financials marked as MDS against a support period marked as MDS, where the financial occurs in the quarter and the support period is active during the quarter
                        string dbFinancialList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                    <entity name='new_palmclientfinancial'>
                                        <attribute name='new_supportperiod' />
                                        <attribute name='new_amount' />
                                        <attribute name='new_mds' />
                                        <attribute name='new_entrydate' />
                                        <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                                <attribute name='new_client' />
                                        </link-entity>
                                        <filter type='and'>
                                            <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                            <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                            <condition entityname='new_palmclientsupport' attribute='new_mds' operator='eq' value='True' />
                                            <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                            <filter type='or'>
                                                <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                                <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                            </filter >
                                        </filter>
                                    </entity>
                                </fetch> ";

                        // Database variables
                        int varClientCount = 0; // Number of clients
                        string dbFirstName = ""; // Firstname field
                        string dbSurname = ""; // Surname field
                        string dbDob = ""; // DOB field
                        string dbDobEst = ""; // DOB flag field - modify to day / month / year fields
                        string dbGender = ""; // Gender field
                        string dbSex = ""; // Sex field
                        string dbCountry = ""; // Country field
                        string dbLanguage = ""; // Language field
                        string dbInterpret = ""; // Interpreter field
                        string dbIndigenous = ""; // Indigenous field
                        string dbClient = ""; // Client number field
                        string dbDepChild = ""; // Dependent children field
                        string dbPalmClientId = ""; // Palm Client Id field
                        string dbCountryCode = ""; // Country Code field
                        string dbLanguageCode = ""; // Language Code field

                        string dbModifiedOn = ""; // Modified On field
                        string dbCarerAvail = ""; // Carer availability field
                        string dbCarerCountry = ""; // Carer country field
                        string dbCarerDob = ""; // Carer DOB field
                        string dbCarerDobFlag = ""; // Carer DOB flag field
                        string dbCarerFirstName = ""; // Carer firstname field
                        string dbCarerIndigenous = ""; // Carer indigenous field
                        string dbCarerLanguage = ""; // Carer language field
                        string dbCarerLocality = ""; // Carer locality field
                        string dbCarerMore = ""; //Carer for more field
                        string dbCarerResidence = ""; // Carer residence field
                        string dbCarerRship = ""; // Carer relationship field
                        string dbCarerSex = ""; // Carer sex field
                        string dbCarerSlk = ""; // Carer SLK field
                        string dbCarerSurname = ""; // Carer surname field
                        string dbCessation = ""; // Cessation field
                        string dbClient2 = ""; // Compare client numebrs field
                        string dbCcpDisType = ""; // Disability type field
                        string dbDva = ""; // DVA field
                        string dbEndDate = ""; // End date field
                        string dbIncomePres = ""; // Income field
                        string dbLivingArrangePres = ""; // Living arrangements field
                        string dbLocality = ""; // Locality field
                        string dbResidentialPres = ""; // Residential field
                        string dbSourceRef = ""; // Source of referral field
                        string dbStartDate = ""; // Start date field
                        string dbCCountryCode = ""; // Carer country code field
                        string dbCLanguageCode = ""; // Carer language code field
                        string dbState = ""; // State field
                        string dbPostcode = ""; // Postcode field
                        string dbCarerState = ""; // Carer state field
                        string dbCarerPostcode = ""; // Carer postcode field

                        // Numeric values for fields
                        int varGender = 0;
                        int varInterpret = 0;
                        int varIndigenous = 0;
                        int varState = 0;
                        int varCarerState = 0;
                        string varClient2 = ""; // support period client number
                        string varSEndDate = ""; // support period end date
                        int varLivingArrange = 0;
                        int varIncome = 0;
                        int varDva = 0;
                        int varResidential = 0;
                        int varCarerSex = 0;
                        int varCarerIndigenous = 0;
                        int varCarerAvail = 0;
                        int varCarerResidence = 0;
                        int varCarerMore = 0;
                        int varCarerRship = 0;
                        int varReferral = 0;
                        int varCessation = 0;
                        int varCCPDisType = 0;
                        int varCCPDepChild = 0;

                        double varCheckTotal = 0; // Check for at least one activity
                        double arrAmount = 0; // Total amount

                        // Totals for activities and financials
                        double varSCPDay = 0;
                        double varSCPNightNon = 0;
                        double varSCPNightActive = 0;
                        double varSCPResidential = 0;
                        double varSCPCounselSupport = 0;

                        double varCCPAssertOut = 0;
                        double varCCPCareCoord = 0;
                        double varCCPHouseAss = 0;
                        double varCCPGroupSocial = 0;

                        double varHSAPAssertOut = 0;
                        double varHSAPCareCoord = 0;
                        double varHSAPHouseAss = 0;

                        double varOPHRAssertOut = 0;
                        double varOPHRCareCoord = 0;
                        double varOPHRHouseAss = 0;
                        double varOPHRGroupSocial = 0;

                        double varSRSCareCoord = 0;
                        double varSRSHouseAss = 0;
                        double varSRSGroupSocial = 0;

                        double varSCPGoodsEquip = 0;
                        double varCCPFlexCare = 0;
                        double varHSAPFlexCare = 0;
                        double varOPHRFlexCare = 0;

                        // Variables for activity data
                        string varAmount = "";
                        string varActivityMDS = "";
                        
                        // Variables for SU drop down list data
                        string varType = "";
                        string varDesc = "";
                        string varMDS = "";

                        string varFn = ""; // SLK first name
                        string varSn = ""; // SLK surname
                        string varSLK = ""; // String to create SLK
                        string varSLKUsed = "."; // Whether SLK has been used before

                        bool varClientSupport = false; // Whether support period has been found
                        string varEntryDate = ""; // Date for activities and financials
                        int varStatusType = 0; // Which support period has been used

                        // Get the fetch XML data and place in entity collection objects
                        EntityCollection result = _service.RetrieveMultiple(new FetchExpression(dbMDSList));
                        EntityCollection result2 = _service.RetrieveMultiple(new FetchExpression(dbDropList));
                        EntityCollection result5 = _service.RetrieveMultiple(new FetchExpression(dbSupportList));
                        EntityCollection result7 = _service.RetrieveMultiple(new FetchExpression(dbActivityList));
                        EntityCollection result8 = _service.RetrieveMultiple(new FetchExpression(dbFinancialList));

                        // Loop through the client data
                        foreach (var c in result.Entities)
                        {
                            //varTest = "STARTING ATTRIBUTES:\r\n";

                            //foreach (KeyValuePair<String, Object> attribute in c.Attributes)
                            //{
                            //    varTest += (attribute.Key + ": " + attribute.Value + "\r\n");
                            //}

                            //varTest += "STARTING FORMATTED:\r\n";

                            //foreach (KeyValuePair<String, String> value in c.FormattedValues)
                            //{
                            //    varTest += (value.Key + ": " + value.Value + "\r\n");
                            //}

                            // Add 1 to client counter
                            varClientCount++;

                            // Process the data as follows:
                            // If there is a formatted value for the field, use it
                            // Otherwise if there is a literal value for the field, use it
                            // Otherwise the value wasn't returned so set as nothing
                            if (c.FormattedValues.Contains("new_firstname"))
                                dbFirstName = c.FormattedValues["new_firstname"];
                            else if(c.Attributes.Contains("new_firstname"))
                                dbFirstName = c.Attributes["new_firstname"].ToString();
                            else
                                dbFirstName = "";

                            if (c.FormattedValues.Contains("new_surname"))
                                dbSurname = c.FormattedValues["new_surname"];
                            else if (c.Attributes.Contains("new_surname"))
                                dbSurname = c.Attributes["new_surname"].ToString();
                            else
                                dbSurname = "";

                            if (c.FormattedValues.Contains("new_dob"))
                                dbDob = c.FormattedValues["new_dob"];
                            else if (c.Attributes.Contains("new_dob"))
                                dbDob = c.Attributes["new_dob"].ToString();
                            else
                                dbDob = "";

                            // Convert date from American format to Australian format
                            dbDob = cleanDateAM(dbDob);

                            if (c.FormattedValues.Contains("new_dobest"))
                                dbDobEst = c.FormattedValues["new_dobest"];
                            else if (c.Attributes.Contains("new_dobest"))
                                dbDobEst = c.Attributes["new_dobest"].ToString();
                            else
                                dbDobEst = "";

                            if (c.FormattedValues.Contains("new_gender"))
                                dbGender = c.FormattedValues["new_gender"];
                            else if (c.Attributes.Contains("new_gender"))
                                dbGender = c.Attributes["new_gender"].ToString();
                            else
                                dbGender = "";

                            if (c.FormattedValues.Contains("new_sex"))
                                dbSex = c.FormattedValues["new_sex"];
                            else if (c.Attributes.Contains("new_sex"))
                                dbSex = c.Attributes["new_sex"].ToString();
                            else
                                dbSex = "";

                            // Sex should be used but historically gender was used
                            // If Sex has a value, replace Gender with that value
                            if (String.IsNullOrEmpty(dbSex) == false)
                                dbGender = dbSex;

                            if (c.FormattedValues.Contains("new_country"))
                                dbCountry = c.FormattedValues["new_country"];
                            else if (c.Attributes.Contains("new_country"))
                                dbCountry = c.Attributes["new_country"].ToString();
                            else
                                dbCountry = "";

                            if (c.FormattedValues.Contains("new_palmddlcountry2.new_code"))
                                dbCountryCode = c.FormattedValues["new_palmddlcountry2.new_code"];
                            else if (c.Attributes.Contains("new_palmddlcountry2.new_code"))
                                dbCountryCode = c.Attributes["new_palmddlcountry2.new_code"].ToString();
                            else
                                dbCountryCode = "";

                            if (c.FormattedValues.Contains("new_language"))
                                dbLanguage = c.FormattedValues["new_language"];
                            else if (c.Attributes.Contains("new_language"))
                                dbLanguage = c.Attributes["new_language"].ToString();
                            else
                                dbLanguage = "";

                            if (c.FormattedValues.Contains("new_palmddllanguage3.new_code"))
                                dbLanguageCode = c.FormattedValues["new_palmddllanguage3.new_code"];
                            else if (c.Attributes.Contains("new_palmddllanguage3.new_code"))
                                dbLanguageCode = c.Attributes["new_palmddllanguage3.new_code"].ToString();
                            else
                                dbLanguageCode = "";

                            if (c.FormattedValues.Contains("new_interpret"))
                                dbInterpret = c.FormattedValues["new_interpret"];
                            else if (c.Attributes.Contains("new_interpret"))
                                dbInterpret = c.Attributes["new_interpret"].ToString();
                            else
                                dbInterpret = "";

                            if (c.FormattedValues.Contains("new_indigenous"))
                                dbIndigenous = c.FormattedValues["new_indigenous"];
                            else if (c.Attributes.Contains("new_indigenous"))
                                dbIndigenous = c.Attributes["new_indigenous"].ToString();
                            else
                                dbIndigenous = "";

                            if (c.FormattedValues.Contains("new_address"))
                                dbClient = c.FormattedValues["new_address"];
                            else if (c.Attributes.Contains("new_address"))
                                dbClient = c.Attributes["new_address"].ToString();
                            else
                                dbClient = "";

                            if (c.FormattedValues.Contains("new_depchild"))
                                dbDepChild = c.FormattedValues["new_depchild"];
                            else if (c.Attributes.Contains("new_depchild"))
                                dbDepChild = c.Attributes["new_depchild"].ToString();
                            else
                                dbDepChild = "";

                            if (c.FormattedValues.Contains("new_palmclientid"))
                                dbPalmClientId = c.FormattedValues["new_palmclientid"];
                            else if (c.Attributes.Contains("new_palmclientid"))
                                dbPalmClientId = c.Attributes["new_palmclientid"].ToString();
                            else
                                dbPalmClientId = "";

                            //Reset support period values
                            dbModifiedOn = "";
                            dbCarerAvail = "";
                            dbCarerCountry = "";
                            dbCarerDob = "";
                            dbCarerDobFlag = "";
                            dbCarerFirstName = "";
                            dbCarerIndigenous = "";
                            dbCarerLanguage = "";
                            dbCarerLocality = "";
                            dbCarerMore = "";
                            dbCarerResidence = "";
                            dbCarerRship = "";
                            dbCarerSex = "";
                            dbCarerSlk = "";
                            dbCarerSurname = "";
                            dbCessation = "";
                            dbClient2 = "";
                            dbCcpDisType = "";
                            dbDva = "";
                            dbEndDate = "";
                            dbIncomePres = "";
                            dbLivingArrangePres = "";
                            dbLocality = "";
                            dbResidentialPres = "";
                            dbSourceRef = "";
                            dbStartDate = "";
                            dbState = "";
                            dbPostcode = "";
                            dbCarerState = "";
                            dbCarerPostcode = "";

                            //Reset support period values
                            varClientSupport = false;
                            varStatusType = 0;

                            // Loop through support period records
                            foreach (var g in result5.Entities)
                            {
                                // We need to get the entity id for the client field for comparisons
                                if (g.Attributes.Contains("new_client"))
                                {
                                    // Get the entity id for the client using the entity reference object
                                    getEntity = (EntityReference)g.Attributes["new_client"];
                                    varClient2 = getEntity.Id.ToString();
                                }
                                else if (g.FormattedValues.Contains("new_client"))
                                    varClient2 = g.FormattedValues["new_client"];
                                else
                                    varClient2 = "";

                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (g.FormattedValues.Contains("new_enddate"))
                                    varSEndDate = g.FormattedValues["new_enddate"];
                                else if (g.Attributes.Contains("new_enddate"))
                                    varSEndDate = g.Attributes["new_enddate"].ToString();
                                else
                                    varSEndDate = "";

                                // Convert date from American format to Australian format
                                varSEndDate = cleanDateAM(varSEndDate);

                                //varTest += dbPalmClientId + " " + varClient2 + "\r\n";

                                // If the client ids are the same then process the support period data
                                // Only 1 support period should be used
                                if (dbPalmClientId == varClient2)
                                {
                                    // Exited support period has lowest preference
                                    if (String.IsNullOrEmpty(varSEndDate) == false && varStatusType < 1)
                                    {
                                        //varTest += "Doing inactive";

                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (g.FormattedValues.Contains("modifiedon"))
                                            dbModifiedOn = g.FormattedValues["modifiedon"];
                                        else if (g.Attributes.Contains("modifiedon"))
                                            dbModifiedOn = g.Attributes["modifiedon"].ToString();
                                        else
                                            dbModifiedOn = "";

                                        // Convert date from American format to Australian format
                                        dbModifiedOn = cleanDateAM(dbModifiedOn);

                                        if (g.FormattedValues.Contains("new_careravail"))
                                            dbCarerAvail = g.FormattedValues["new_careravail"];
                                        else if (g.Attributes.Contains("new_careravail"))
                                            dbCarerAvail = g.Attributes["new_careravail"].ToString();
                                        else
                                            dbCarerAvail = "";

                                        if (g.FormattedValues.Contains("new_carercountry"))
                                            dbCarerCountry = g.FormattedValues["new_carercountry"];
                                        else if (g.Attributes.Contains("new_carercountry"))
                                            dbCarerCountry = g.Attributes["new_carercountry"].ToString();
                                        else
                                            dbCarerCountry = "";

                                        if (g.FormattedValues.Contains("new_palmddlcountry2.new_code"))
                                            dbCCountryCode = g.FormattedValues["new_palmddlcountry2.new_code"];
                                        else if (g.Attributes.Contains("new_palmddlcountry2.new_code"))
                                            dbCCountryCode = g.Attributes["new_palmddlcountry2.new_code"].ToString();
                                        else
                                            dbCCountryCode = "";

                                        if (g.FormattedValues.Contains("new_carerdob"))
                                            dbCarerDob = g.FormattedValues["new_carerdob"];
                                        else if (g.Attributes.Contains("new_carerdob"))
                                            dbCarerDob = g.Attributes["new_carerdob"].ToString();
                                        else
                                            dbCarerDob = "";

                                        // Convert date from American format to Australian format
                                        dbCarerDob = cleanDateAM(dbCarerDob);

                                        if (g.FormattedValues.Contains("new_carerdobflag"))
                                            dbCarerDobFlag = g.FormattedValues["new_carerdobflag"];
                                        else if (g.Attributes.Contains("new_carerdobflag"))
                                            dbCarerDobFlag = g.Attributes["new_carerdobflag"].ToString();
                                        else
                                            dbCarerDobFlag = "";

                                        if (g.FormattedValues.Contains("new_carerfirstname"))
                                            dbCarerFirstName = g.FormattedValues["new_carerfirstname"];
                                        else if (g.Attributes.Contains("new_carerfirstname"))
                                            dbCarerFirstName = g.Attributes["new_carerfirstname"].ToString();
                                        else
                                            dbCarerFirstName = "";

                                        if (g.FormattedValues.Contains("new_carerindigenous"))
                                            dbCarerIndigenous = g.FormattedValues["new_carerindigenous"];
                                        else if (g.Attributes.Contains("new_carerindigenous"))
                                            dbCarerIndigenous = g.Attributes["new_carerindigenous"].ToString();
                                        else
                                            dbCarerIndigenous = "";

                                        if (g.FormattedValues.Contains("new_carerlanguage"))
                                            dbCarerLanguage = g.FormattedValues["new_carerlanguage"];
                                        else if (g.Attributes.Contains("new_carerlanguage"))
                                            dbCarerLanguage = g.Attributes["new_carerlanguage"].ToString();
                                        else
                                            dbCarerLanguage = "";

                                        if (g.FormattedValues.Contains("new_palmddllanguage3.new_code"))
                                            dbCLanguageCode = g.FormattedValues["new_palmddllanguage3.new_code"];
                                        else if (g.Attributes.Contains("new_palmddllanguage3.new_code"))
                                            dbCLanguageCode = g.Attributes["new_palmddllanguage3.new_code"].ToString();
                                        else
                                            dbCLanguageCode = "";

                                        if (g.FormattedValues.Contains("new_carerlocality"))
                                            dbCarerLocality = g.FormattedValues["new_carerlocality"];
                                        else if (g.Attributes.Contains("new_carerlocality"))
                                            dbCarerLocality = g.Attributes["new_carerlocality"].ToString();
                                        else
                                            dbCarerLocality = "";

                                        if (g.FormattedValues.Contains("new_palmddllocality5.new_state"))
                                            dbCarerState = g.FormattedValues["new_palmddllocality5.new_state"];
                                        else if (g.Attributes.Contains("new_palmddllocality5.new_state"))
                                            dbCarerState = g.Attributes["new_palmddllocality5.new_state"].ToString();
                                        else
                                            dbCarerState = "";

                                        if (g.FormattedValues.Contains("new_palmddllocality5.new_postcode"))
                                            dbCarerPostcode = g.FormattedValues["new_palmddllocality5.new_postcode"];
                                        else if (g.Attributes.Contains("new_palmddllocality5.new_postcode"))
                                            dbCarerPostcode = g.Attributes["new_palmddllocality5.new_postcode"].ToString();
                                        else
                                            dbCarerPostcode = "";

                                        if (g.FormattedValues.Contains("new_carermore"))
                                            dbCarerMore = g.FormattedValues["new_carermore"];
                                        else if (g.Attributes.Contains("new_carermore"))
                                            dbCarerMore = g.Attributes["new_carermore"].ToString();
                                        else
                                            dbCarerMore = "";

                                        if (g.FormattedValues.Contains("new_carerresidence"))
                                            dbCarerResidence = g.FormattedValues["new_carerresidence"];
                                        else if (g.Attributes.Contains("new_carerresidence"))
                                            dbCarerResidence = g.Attributes["new_carerresidence"].ToString();
                                        else
                                            dbCarerResidence = "";

                                        if (g.FormattedValues.Contains("new_carerrship"))
                                            dbCarerRship = g.FormattedValues["new_carerrship"];
                                        else if (g.Attributes.Contains("new_carerrship"))
                                            dbCarerRship = g.Attributes["new_carerrship"].ToString();
                                        else
                                            dbCarerRship = "";

                                        if (g.FormattedValues.Contains("new_carersex"))
                                            dbCarerSex = g.FormattedValues["new_carersex"];
                                        else if (g.Attributes.Contains("new_carersex"))
                                            dbCarerSex = g.Attributes["new_carersex"].ToString();
                                        else
                                            dbCarerSex = "";

                                        if (g.FormattedValues.Contains("new_carersurname"))
                                            dbCarerSurname = g.FormattedValues["new_carersurname"];
                                        else if (g.Attributes.Contains("new_carersurname"))
                                            dbCarerSurname = g.Attributes["new_carersurname"].ToString();
                                        else
                                            dbCarerSurname = "";

                                        if (g.FormattedValues.Contains("new_cessation"))
                                            dbCessation = g.FormattedValues["new_cessation"];
                                        else if (g.Attributes.Contains("new_cessation"))
                                            dbCessation = g.Attributes["new_cessation"].ToString();
                                        else
                                            dbCessation = "";

                                        if (g.FormattedValues.Contains("new_client"))
                                            dbClient2 = g.FormattedValues["new_client"];
                                        else if (g.Attributes.Contains("new_client"))
                                            dbClient2 = g.Attributes["new_client"].ToString();
                                        else
                                            dbClient2 = "";

                                        if (g.FormattedValues.Contains("new_ccpdistype"))
                                            dbCcpDisType = g.FormattedValues["new_ccpdistype"];
                                        else if (g.Attributes.Contains("new_ccpdistype"))
                                            dbCcpDisType = g.Attributes["new_ccpdistype"].ToString();
                                        else
                                            dbCcpDisType = "";

                                        if (g.FormattedValues.Contains("new_dva"))
                                            dbDva = g.FormattedValues["new_dva"];
                                        else if (g.Attributes.Contains("new_dva"))
                                            dbDva = g.Attributes["new_dva"].ToString();
                                        else
                                            dbDva = "";

                                        if (g.FormattedValues.Contains("new_enddate"))
                                            dbEndDate = g.FormattedValues["new_enddate"];
                                        else if (g.Attributes.Contains("new_enddate"))
                                            dbEndDate = g.Attributes["new_enddate"].ToString();
                                        else
                                            dbEndDate = "";

                                        // Convert date from American format to Australian format
                                        dbEndDate = cleanDateAM(dbEndDate);

                                        if (g.FormattedValues.Contains("new_incomepres"))
                                            dbIncomePres = g.FormattedValues["new_incomepres"];
                                        else if (g.Attributes.Contains("new_incomepres"))
                                            dbIncomePres = g.Attributes["new_incomepres"].ToString();
                                        else
                                            dbIncomePres = "";

                                        if (g.FormattedValues.Contains("new_livingarrangepres"))
                                            dbLivingArrangePres = g.FormattedValues["new_livingarrangepres"];
                                        else if (g.Attributes.Contains("new_livingarrangepres"))
                                            dbLivingArrangePres = g.Attributes["new_livingarrangepres"].ToString();
                                        else
                                            dbLivingArrangePres = "";

                                        if (g.FormattedValues.Contains("new_locality"))
                                            dbLocality = g.FormattedValues["new_locality"];
                                        else if (g.Attributes.Contains("new_locality"))
                                            dbLocality = g.Attributes["new_locality"].ToString();
                                        else
                                            dbLocality = "";

                                        if (g.FormattedValues.Contains("new_palmddllocality4.new_state"))
                                            dbState = g.FormattedValues["new_palmddllocality4.new_state"];
                                        else if (g.Attributes.Contains("new_palmddllocality4.new_state"))
                                            dbState = g.Attributes["new_palmddllocality4.new_state"].ToString();
                                        else
                                            dbState = "";

                                        if (g.FormattedValues.Contains("new_palmddllocality4.new_postcode"))
                                            dbPostcode = g.FormattedValues["new_palmddllocality4.new_postcode"];
                                        else if (g.Attributes.Contains("new_palmddllocality4.new_postcode"))
                                            dbPostcode = g.Attributes["new_palmddllocality4.new_postcode"].ToString();
                                        else
                                            dbPostcode = "";

                                        if (g.FormattedValues.Contains("new_residentialpres"))
                                            dbResidentialPres = g.FormattedValues["new_residentialpres"];
                                        else if (g.Attributes.Contains("new_residentialpres"))
                                            dbResidentialPres = g.Attributes["new_residentialpres"].ToString();
                                        else
                                            dbResidentialPres = "";

                                        if (g.FormattedValues.Contains("new_sourceref"))
                                            dbSourceRef = g.FormattedValues["new_sourceref"];
                                        else if (g.Attributes.Contains("new_sourceref"))
                                            dbSourceRef = g.Attributes["new_sourceref"].ToString();
                                        else
                                            dbSourceRef = "";

                                        if (g.FormattedValues.Contains("new_startdate"))
                                            dbStartDate = g.FormattedValues["new_startdate"];
                                        else if (g.Attributes.Contains("new_startdate"))
                                            dbStartDate = g.Attributes["new_startdate"].ToString();
                                        else
                                            dbStartDate = "";

                                        // Convert date from American format to Australian format
                                        dbStartDate = cleanDateAM(dbStartDate);

                                        // Remove invalid firstname and surname characters and add to SLK with trialing characters
                                        varFn = cleanString(dbCarerFirstName.ToUpper(), "slk") + "222";
                                        varSn = cleanString(dbCarerSurname.ToUpper(), "slk") + "22222";

                                        // Create carer SLK
                                        dbCarerSlk = varSn.Substring(1, 1) + varSn.Substring(2, 1) + varSn.Substring(4, 1) + varFn.Substring(1, 1) + varFn.Substring(2, 1);

                                        // Set as null if there were no characters
                                        if (dbCarerSlk == "22222")
                                            dbCarerSlk = "";

                                        varStatusType = 1; // Inactive support period used
                                        varClientSupport = true; // Support period found
                                    }
                                    //Active support period has highest preference
                                    else if (String.IsNullOrEmpty(varSEndDate) == true && varStatusType < 2)
                                    {
                                        //varTest += "Doing active";

                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (g.FormattedValues.Contains("modifiedon"))
                                            dbModifiedOn = g.FormattedValues["modifiedon"];
                                        else if (g.Attributes.Contains("modifiedon"))
                                            dbModifiedOn = g.Attributes["modifiedon"].ToString();
                                        else
                                            dbModifiedOn = "";

                                        // Convert date from American format to Australian format
                                        dbModifiedOn = cleanDateAM(dbModifiedOn);

                                        if (g.FormattedValues.Contains("new_careravail"))
                                            dbCarerAvail = g.FormattedValues["new_careravail"];
                                        else if (g.Attributes.Contains("new_careravail"))
                                            dbCarerAvail = g.Attributes["new_careravail"].ToString();
                                        else
                                            dbCarerAvail = "";

                                        if (g.FormattedValues.Contains("new_carercountry"))
                                            dbCarerCountry = g.FormattedValues["new_carercountry"];
                                        else if (g.Attributes.Contains("new_carercountry"))
                                            dbCarerCountry = g.Attributes["new_carercountry"].ToString();
                                        else
                                            dbCarerCountry = "";

                                        if (g.FormattedValues.Contains("new_palmddlcountry2.new_code"))
                                            dbCCountryCode = g.FormattedValues["new_palmddlcountry2.new_code"];
                                        else if (g.Attributes.Contains("new_palmddlcountry2.new_code"))
                                            dbCCountryCode = g.Attributes["new_palmddlcountry2.new_code"].ToString();
                                        else
                                            dbCCountryCode = "";

                                        if (g.FormattedValues.Contains("new_carerdob"))
                                            dbCarerDob = g.FormattedValues["new_carerdob"];
                                        else if (g.Attributes.Contains("new_carerdob"))
                                            dbCarerDob = g.Attributes["new_carerdob"].ToString();
                                        else
                                            dbCarerDob = "";

                                        // Convert date from American format to Australian format
                                        dbCarerDob = cleanDateAM(dbCarerDob);

                                        if (g.FormattedValues.Contains("new_carerdobflag"))
                                            dbCarerDobFlag = g.FormattedValues["new_carerdobflag"];
                                        else if (g.Attributes.Contains("new_carerdobflag"))
                                            dbCarerDobFlag = g.Attributes["new_carerdobflag"].ToString();
                                        else
                                            dbCarerDobFlag = "";

                                        if (g.FormattedValues.Contains("new_carerfirstname"))
                                            dbCarerFirstName = g.FormattedValues["new_carerfirstname"];
                                        else if (g.Attributes.Contains("new_carerfirstname"))
                                            dbCarerFirstName = g.Attributes["new_carerfirstname"].ToString();
                                        else
                                            dbCarerFirstName = "";

                                        if (g.FormattedValues.Contains("new_carerindigenous"))
                                            dbCarerIndigenous = g.FormattedValues["new_carerindigenous"];
                                        else if (g.Attributes.Contains("new_carerindigenous"))
                                            dbCarerIndigenous = g.Attributes["new_carerindigenous"].ToString();
                                        else
                                            dbCarerIndigenous = "";

                                        if (g.FormattedValues.Contains("new_carerlanguage"))
                                            dbCarerLanguage = g.FormattedValues["new_carerlanguage"];
                                        else if (g.Attributes.Contains("new_carerlanguage"))
                                            dbCarerLanguage = g.Attributes["new_carerlanguage"].ToString();
                                        else
                                            dbCarerLanguage = "";

                                        if (g.FormattedValues.Contains("new_palmddllanguage3.new_code"))
                                            dbCLanguageCode = g.FormattedValues["new_palmddllanguage3.new_code"];
                                        else if (g.Attributes.Contains("new_palmddllanguage3.new_code"))
                                            dbCLanguageCode = g.Attributes["new_palmddllanguage3.new_code"].ToString();
                                        else
                                            dbCLanguageCode = "";

                                        if (g.FormattedValues.Contains("new_carerlocality"))
                                            dbCarerLocality = g.FormattedValues["new_carerlocality"];
                                        else if (g.Attributes.Contains("new_carerlocality"))
                                            dbCarerLocality = g.Attributes["new_carerlocality"].ToString();
                                        else
                                            dbCarerLocality = "";

                                        if (g.FormattedValues.Contains("new_palmddllocality5.new_state"))
                                            dbCarerState = g.FormattedValues["new_palmddllocality5.new_state"];
                                        else if (g.Attributes.Contains("new_palmddllocality5.new_state"))
                                            dbCarerState = g.Attributes["new_palmddllocality5.new_state"].ToString();
                                        else
                                            dbCarerState = "";

                                        if (g.FormattedValues.Contains("new_palmddllocality5.new_postcode"))
                                            dbCarerPostcode = g.FormattedValues["new_palmddllocality5.new_postcode"];
                                        else if (g.Attributes.Contains("new_palmddllocality5.new_postcode"))
                                            dbCarerPostcode = g.Attributes["new_palmddllocality5.new_postcode"].ToString();
                                        else
                                            dbCarerPostcode = "";

                                        if (g.FormattedValues.Contains("new_carermore"))
                                            dbCarerMore = g.FormattedValues["new_carermore"];
                                        else if (g.Attributes.Contains("new_carermore"))
                                            dbCarerMore = g.Attributes["new_carermore"].ToString();
                                        else
                                            dbCarerMore = "";

                                        if (g.FormattedValues.Contains("new_carerresidence"))
                                            dbCarerResidence = g.FormattedValues["new_carerresidence"];
                                        else if (g.Attributes.Contains("new_carerresidence"))
                                            dbCarerResidence = g.Attributes["new_carerresidence"].ToString();
                                        else
                                            dbCarerResidence = "";

                                        if (g.FormattedValues.Contains("new_carerrship"))
                                            dbCarerRship = g.FormattedValues["new_carerrship"];
                                        else if (g.Attributes.Contains("new_carerrship"))
                                            dbCarerRship = g.Attributes["new_carerrship"].ToString();
                                        else
                                            dbCarerRship = "";

                                        if (g.FormattedValues.Contains("new_carersex"))
                                            dbCarerSex = g.FormattedValues["new_carersex"];
                                        else if (g.Attributes.Contains("new_carersex"))
                                            dbCarerSex = g.Attributes["new_carersex"].ToString();
                                        else
                                            dbCarerSex = "";

                                        if (g.FormattedValues.Contains("new_carersurname"))
                                            dbCarerSurname = g.FormattedValues["new_carersurname"];
                                        else if (g.Attributes.Contains("new_carersurname"))
                                            dbCarerSurname = g.Attributes["new_carersurname"].ToString();
                                        else
                                            dbCarerSurname = "";

                                        if (g.FormattedValues.Contains("new_cessation"))
                                            dbCessation = g.FormattedValues["new_cessation"];
                                        else if (g.Attributes.Contains("new_cessation"))
                                            dbCessation = g.Attributes["new_cessation"].ToString();
                                        else
                                            dbCessation = "";

                                        if (g.FormattedValues.Contains("new_client"))
                                            dbClient2 = g.FormattedValues["new_client"];
                                        else if (g.Attributes.Contains("new_client"))
                                            dbClient2 = g.Attributes["new_client"].ToString();
                                        else
                                            dbClient2 = "";

                                        if (g.FormattedValues.Contains("new_ccpdistype"))
                                            dbCcpDisType = g.FormattedValues["new_ccpdistype"];
                                        else if (g.Attributes.Contains("new_ccpdistype"))
                                            dbCcpDisType = g.Attributes["new_ccpdistype"].ToString();
                                        else
                                            dbCcpDisType = "";

                                        if (g.FormattedValues.Contains("new_dva"))
                                            dbDva = g.FormattedValues["new_dva"];
                                        else if (g.Attributes.Contains("new_dva"))
                                            dbDva = g.Attributes["new_dva"].ToString();
                                        else
                                            dbDva = "";

                                        if (g.FormattedValues.Contains("new_enddate"))
                                            dbEndDate = g.FormattedValues["new_enddate"];
                                        else if (g.Attributes.Contains("new_enddate"))
                                            dbEndDate = g.Attributes["new_enddate"].ToString();
                                        else
                                            dbEndDate = "";

                                        // Convert date from American format to Australian format
                                        dbEndDate = cleanDateAM(dbEndDate);


                                        if (g.FormattedValues.Contains("new_incomepres"))
                                            dbIncomePres = g.FormattedValues["new_incomepres"];
                                        else if (g.Attributes.Contains("new_incomepres"))
                                            dbIncomePres = g.Attributes["new_incomepres"].ToString();
                                        else
                                            dbIncomePres = "";

                                        if (g.FormattedValues.Contains("new_livingarrangepres"))
                                            dbLivingArrangePres = g.FormattedValues["new_livingarrangepres"];
                                        else if (g.Attributes.Contains("new_livingarrangepres"))
                                            dbLivingArrangePres = g.Attributes["new_livingarrangepres"].ToString();
                                        else
                                            dbLivingArrangePres = "";

                                        if (g.FormattedValues.Contains("new_locality"))
                                            dbLocality = g.FormattedValues["new_locality"];
                                        else if (g.Attributes.Contains("new_locality"))
                                            dbLocality = g.Attributes["new_locality"].ToString();
                                        else
                                            dbLocality = "";

                                        if (g.FormattedValues.Contains("new_palmddllocality4.new_state"))
                                            dbState = g.FormattedValues["new_palmddllocality4.new_state"];
                                        else if (g.Attributes.Contains("new_palmddllocality4.new_state"))
                                            dbState = g.Attributes["new_palmddllocality4.new_state"].ToString();
                                        else
                                            dbState = "";

                                        if (g.FormattedValues.Contains("new_palmddllocality4.new_postcode"))
                                            dbPostcode = g.FormattedValues["new_palmddllocality4.new_postcode"];
                                        else if (g.Attributes.Contains("new_palmddllocality4.new_postcode"))
                                            dbPostcode = g.Attributes["new_palmddllocality4.new_postcode"].ToString();
                                        else
                                            dbPostcode = "";

                                        if (g.FormattedValues.Contains("new_residentialpres"))
                                            dbResidentialPres = g.FormattedValues["new_residentialpres"];
                                        else if (g.Attributes.Contains("new_residentialpres"))
                                            dbResidentialPres = g.Attributes["new_residentialpres"].ToString();
                                        else
                                            dbResidentialPres = "";

                                        if (g.FormattedValues.Contains("new_sourceref"))
                                            dbSourceRef = g.FormattedValues["new_sourceref"];
                                        else if (g.Attributes.Contains("new_sourceref"))
                                            dbSourceRef = g.Attributes["new_sourceref"].ToString();
                                        else
                                            dbSourceRef = "";

                                        if (g.FormattedValues.Contains("new_startdate"))
                                            dbStartDate = g.FormattedValues["new_startdate"];
                                        else if (g.Attributes.Contains("new_startdate"))
                                            dbStartDate = g.Attributes["new_startdate"].ToString();
                                        else
                                            dbStartDate = "";

                                        // Convert date from American format to Australian format
                                        dbStartDate = cleanDateAM(dbStartDate);

                                        // Remove invalid firstname and surname characters and add to SLK with trialing characters
                                        varFn = cleanString(dbCarerFirstName.ToUpper(), "slk") + "222";
                                        varSn = cleanString(dbCarerSurname.ToUpper(), "slk") + "22222";

                                        // Create carer SLK
                                        dbCarerSlk = varSn.Substring(1, 1) + varSn.Substring(2, 1) + varSn.Substring(4, 1) + varFn.Substring(1, 1) + varFn.Substring(2, 1);

                                        // Set as null if there were no characters
                                        if (dbCarerSlk == "22222")
                                            dbCarerSlk = "";

                                        varStatusType = 2; // Active support period used
                                        varClientSupport = true; // Support period found
                                        break; // Active support period found - exit loop
                                    }
                                } // Same client
                            } // Support period loop
                            
                            // Reset numeric values
                            varGender = 0;
                            varInterpret = 0;
                            varIndigenous = 0;
                            varState = 0;
                            varCarerState = 0;
                            varLivingArrange = 0;
                            varIncome = 0;
                            varDva = 0;
                            varResidential = 0;
                            varCarerSex = 0;
                            varCarerIndigenous = 0;
                            varCarerAvail = 0;
                            varCarerResidence = 0;
                            varCarerMore = 0;
                            varCarerRship = 0;
                            varReferral = 0;
                            varCessation = 0;
                            varCCPDisType = 0;
                            varCCPDepChild = 0;

                            // Create country code if don't know or not applicable
                            if (dbCountry == "Don't Know" || dbCountry == "Not Applicable")
                                dbCountryCode = "9999";
                            if (dbCarerCountry == "Don't Know" || dbCarerCountry == "Not Applicable")
                                dbCCountryCode = "9999";

                            // Loop through SU Drop Down List values to get numeric values
                            foreach (var d in result2.Entities)
                            {
                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (d.FormattedValues.Contains("new_type"))
                                    varType = d.FormattedValues["new_type"];
                                else if (d.Attributes.Contains("new_type"))
                                    varType = d.Attributes["new_type"].ToString();
                                else
                                    varType = "";

                                if (d.FormattedValues.Contains("new_description"))
                                    varDesc = d.FormattedValues["new_description"];
                                else if (d.Attributes.Contains("new_description"))
                                    varDesc = d.Attributes["new_description"].ToString();
                                else
                                    varDesc = "";

                                if (d.FormattedValues.Contains("new_mds"))
                                    varMDS = d.FormattedValues["new_mds"];
                                else if (d.Attributes.Contains("new_mds"))
                                    varMDS = d.Attributes["new_mds"].ToString();
                                else
                                    varMDS = "";

                                // Make sure MDS value is numeric, or set to 0
                                varMDS = cleanString(varMDS, "number");
                                if (String.IsNullOrEmpty(varMDS) == true)
                                    varMDS = "0";

                                // If the drop down type is sex, compare with sex fields
                                if (varType.ToLower() == "sex")
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    if (dbGender.ToLower() == varDesc.ToLower())
                                        Int32.TryParse(varMDS, out varGender);
                                    if (dbCarerSex.ToLower() == varDesc.ToLower())
                                        Int32.TryParse(varMDS, out varCarerSex);
                                }
                                // If the drop down type is interpret, compare with need for intepreter field
                                if (varType.ToLower() == "interpret" && dbInterpret.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varInterpret);
                                }
                                // If the drop down type is indigenous, compare with indigenous status field
                                if (varType.ToLower() == "indigenous")
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    if (dbIndigenous.ToLower() == varDesc.ToLower())
                                        Int32.TryParse(varMDS, out varIndigenous);
                                    if (dbCarerIndigenous.ToLower() == varDesc.ToLower())
                                        Int32.TryParse(varMDS, out varCarerIndigenous);
                                }
                                // If the drop down type is state, compare with state fields
                                if (varType.ToLower() == "state")
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    if (dbState.ToLower() == varDesc.ToLower())
                                        Int32.TryParse(varMDS, out varState);
                                    if (dbCarerState.ToLower() == varDesc.ToLower())
                                        Int32.TryParse(varMDS, out varCarerState);
                                }
                                // If the drop down type is livingarrange, compare with living arrangements field
                                if (varType.ToLower() == "livingarrange" && dbLivingArrangePres.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varLivingArrange);
                                }
                                // If the drop down type is income, compare with income field
                                if (varType.ToLower() == "income" && dbIncomePres.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varIncome);
                                }
                                // If the drop down type is dva, compare with dva field
                                if (varType.ToLower() == "dva" && dbDva.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varDva);
                                }
                                // If the drop down type is residential, compare with residential field
                                if (varType.ToLower() == "residential" && dbResidentialPres.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varResidential);
                                }
                                // If the drop down type is careravail, compare with carer avaiability field
                                if (varType.ToLower() == "careravail" && dbCarerAvail.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varCarerAvail);
                                }
                                // If the drop down type is carerresidence, compare with carer residence field
                                if (varType.ToLower() == "carerresidence" && dbCarerResidence.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varCarerResidence);
                                }
                                // If the drop down type is carermore, compare with carer for more field
                                if (varType.ToLower() == "carermore" && dbCarerMore.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varCarerMore);
                                }
                                // If the drop down type is crship, compare with carer relationship field
                                if (varType.ToLower() == "crship" && dbCarerRship.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varCarerRship);
                                }
                                // If the drop down type is sourceref, compare with source of referral field
                                if (varType.ToLower() == "sourceref" && dbSourceRef.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varReferral);
                                }
                                // If the drop down type is ccpdepchild, compare with dependent children field
                                if (varType.ToLower() == "ccpdepchild" && dbDepChild.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varCCPDepChild);
                                }
                                // If the drop down type is ccpdis, compare with disability field
                                if (varType.ToLower() == "ccpdis" && dbCcpDisType.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varCCPDisType);
                                }
                                // If the drop down type is cessation, compare with cessation field
                                if (varType.ToLower() == "cessation" && dbCessation.ToLower() == varDesc.ToLower())
                                {
                                    // If the string values match, set the numeric value to the MDS value from the Drop Down List
                                    Int32.TryParse(varMDS, out varCessation);
                                }
                            }

                            //varTest += dbFirstName + " " + dbSurname + " " + varClientSupport + " " + dbStartDate + " ." + dbLocality + ".\r\n";

                            // Begin to create client part of extract
                            sbClientList.Append("HACC,");
                            sbClientList.Append("201,");
                            sbClientList.Append("CLIENT,");

                            // Remove invalid firstname and surname characters and add to SLK with trialing characters
                            varFn = cleanString(dbFirstName.ToUpper(), "slk") + "222";
                            varSn = cleanString(dbSurname.ToUpper(), "slk") + "22222";

                            // Create client SLK
                            varSLK = varSn.Substring(1, 1) + varSn.Substring(2, 1) + varSn.Substring(4, 1) + varFn.Substring(1, 1) + varFn.Substring(2, 1);

                            // Set as null if there were no characters
                            if (varSLK == "22222")
                                varSLK = "";

                            // If the SLK has already been used, change the last character to a 2
                            // Append the SLK to the SLK used string
                            if (String.IsNullOrEmpty(varSLK) == false)
                            {
                                if (varSLKUsed.IndexOf(varSLK) > -1)
                                {
                                    varSLK = varSLK.Substring(0, 4) + "2";
                                    varSLKUsed += "*" + varSLK + "*";
                                }
                                else
                                    varSLKUsed += "*" + varSLK + "*";
                            }

                            // Add SLK to extract or alert of error
                            if (String.IsNullOrEmpty(varSLK) == false && varSLK.Length == 5)
                                sbClientList.Append(varSLK + ",");
                            else
                            {
                                sbClientList.Append(",");
                                sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - No SLK");
                            }

                            // Add date of birth or advise of error
                            if (String.IsNullOrEmpty(dbDob) == false)
                            {
                                // Date of birth needs to be greater than 1895
                                if (Convert.ToDateTime(dbDob).Year > 1895)
                                    sbClientList.Append(cleanDateM(Convert.ToDateTime(dbDob)) + ",");
                                else
                                {
                                    sbClientList.Append(",");
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - DOB");
                                }

                            }
                            else
                            {
                                // If there is no date of birth create a default one
                                sbClientList.Append("01/01/1950,");
                                // If the date of birth was not estimated / unknown then alert of error
                                if (dbDobEst != "False")
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - DOB");
                            }

                            //Append dob flag or advise of error
                            if (dbDobEst == "True" || dbDobEst == "Yes")
                                sbClientList.Append("1,");
                            else if (dbDobEst == "False" || dbDobEst == "No")
                                sbClientList.Append("2,");
                            else
                            {
                                sbClientList.Append(",");
                                sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - DOB Estimate Flag");
                            }

                            // Append gender or advise of error
                            if (varGender > 0)
                                sbClientList.Append(varGender + ",");
                            else
                            {
                                // Append default value
                                sbClientList.Append("9,");
                                // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                if (String.IsNullOrEmpty(dbGender) == false)
                                    sbErrorList.AppendLine("Invalid Gender Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbGender + "");
                                else
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client Gender");
                            }

                            //Append country and language or alert of missing data
                            if (String.IsNullOrEmpty(dbCountryCode) == false)
                                sbClientList.Append(dbCountryCode + ",");
                            else
                            {
                                sbClientList.Append("9999,");
                                sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client Country");
                            }
                            if (String.IsNullOrEmpty(dbLanguageCode) == false)
                                sbClientList.Append(dbLanguageCode + ",");
                            else
                            {
                                sbClientList.Append("9999,");
                                sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client Language");
                            }

                            // Append need for interpreter
                            if (varInterpret > 0)
                                sbClientList.Append(varInterpret + ",");
                            else
                            {
                                // Append default value
                                sbClientList.Append("9,");
                                // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                if (String.IsNullOrEmpty(dbInterpret) == false)
                                    sbErrorList.AppendLine("Invalid Interpreter Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbInterpret + "");
                                else
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client Need for Intepreter");
                            }

                            // Append indigenous status
                            if (varIndigenous > 0)
                                sbClientList.Append(varIndigenous + ",");
                            else
                            {
                                // Append default value
                                sbClientList.Append("99,");
                                // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                if (String.IsNullOrEmpty(dbIndigenous) == false)
                                    sbErrorList.AppendLine("Invalid Indigenous Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbIndigenous + "");
                                else
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client Indigenous State");
                            }

                            // Append state
                            if (varState > 0)
                                sbClientList.Append(varState + ",");
                            else
                            {
                                // Append default value
                                sbClientList.Append("9,");
                                // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                if (String.IsNullOrEmpty(dbState) == false)
                                    sbErrorList.AppendLine("Invalid State Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbState + "");
                                else
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client State");
                            }

                            // Append locality or advise of missing data
                            if (String.IsNullOrEmpty(dbLocality) == false)
                            {
                                // Locality must be less than 46 characters
                                if (dbLocality.Length <= 46)
                                    sbClientList.Append(dbLocality + ",");
                                else
                                {
                                    sbClientList.Append(",");
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client Locality");
                                }
                            }
                            else
                            {
                                sbClientList.Append(",");
                                sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client Locality");
                            }

                            // Append postcode if numeric or alert of error
                            if (Int32.TryParse(dbPostcode, out varCheckInt))
                                sbClientList.Append(dbPostcode + ",");
                            else
                            {
                                sbClientList.Append("9999,");
                                sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client Postcode");
                            }

                            //varTest += dbCarerAvail + "\r\n";

                            // Append SLK flag based on carer availability (hardcoded)
                            if (dbCarerAvail.ToLower() == "has a carer")
                                sbClientList.Append("2,");
                            else if (dbCarerAvail.ToLower() == "has no carer")
                                sbClientList.Append("1,");
                            else
                            {
                                sbClientList.Append("1,");
                                sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - SLK Missing Flag");
                            }

                            // Append living arrangements or alert of error
                            if (varLivingArrange > 0)
                                sbClientList.Append(varLivingArrange + ",");
                            else
                            {
                                // Append default value
                                sbClientList.Append("9,");
                                // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                if (String.IsNullOrEmpty(dbLivingArrangePres) == false)
                                    sbErrorList.AppendLine("Invalid Living Arrangements Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbLivingArrangePres + "");
                                else
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client Living Arrangements");
                            }

                            // Append income or alert of error
                            if (varIncome > 0)
                                sbClientList.Append(varIncome + ",");
                            else
                            {
                                // Append default value
                                sbClientList.Append("9,");
                                // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                if (String.IsNullOrEmpty(dbIncomePres) == false)
                                    sbErrorList.AppendLine("Invalid Income Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbIncomePres + "");
                                else
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client Income");
                            }

                            // Append DVA or alert of error
                            if (varDva > 0)
                                sbClientList.Append(varDva + ",");
                            else
                            {
                                // Append default value
                                sbClientList.Append("9,");
                                // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                if (String.IsNullOrEmpty(dbDva) == false)
                                    sbErrorList.AppendLine("Invalid DVA Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbDva + "");
                                else
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client DVA");
                            }

                            // Append Residential or alert of error
                            if (varResidential > 0)
                                sbClientList.Append(varResidential + ",");
                            else
                            {
                                // Append default value
                                sbClientList.Append("99,");
                                // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                if (String.IsNullOrEmpty(dbResidentialPres) == false)
                                    sbErrorList.AppendLine("Invalid Residential Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbResidentialPres + "");
                                else
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client Residential");
                            }

                            // Append Carer Availability (hardcoded)
                            if (dbCarerAvail.ToLower() == "has a carer")
                                sbClientList.Append("1,");
                            else if (dbCarerAvail.ToLower() == "has no carer")
                                sbClientList.Append("2,");
                            else
                            {
                                sbClientList.Append("2,");
                                sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer availability");
                            }

                            //Carer information is only appended if there is a carer
                            if (dbCarerAvail.ToLower() == "has a carer")
                            {
                                // Append Carer SLK
                                if (String.IsNullOrEmpty(dbCarerSlk) == false)
                                    sbClientList.Append(dbCarerSlk + ",");
                                else
                                    sbClientList.Append(",");

                                // Append Carer DOB
                                if (String.IsNullOrEmpty(dbCarerDob) == false)
                                {
                                    // DOB must be greater than 1895 or alert of error
                                    if (Convert.ToDateTime(dbCarerDob).Year > 1895)
                                    {
                                        sbClientList.Append(cleanDateM(Convert.ToDateTime(dbCarerDob)) + ",");
                                    }
                                    else
                                    {
                                        sbClientList.Append(",");
                                        sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer DOB");
                                    }

                                }
                                else
                                {
                                    // Default date if missing
                                    sbClientList.Append("01/01/1950,");
                                    // Error if DOB missing but not estimated or unknown
                                    if (dbCarerDobFlag != "False")
                                        sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer DOB");
                                }

                                //Append dob flag or advise of error
                                if (dbCarerDobFlag == "True" || dbCarerDobFlag == "Yes")
                                    sbClientList.Append("1,");
                                else if (dbCarerDobFlag == "False" || dbCarerDobFlag == "No")
                                    sbClientList.Append("2,");
                                else
                                {
                                    sbClientList.Append(",");
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer DOB Estimate Flag");
                                }

                                // Append carer sex or advise of error
                                if (varCarerSex > 0)
                                    sbClientList.Append(varCarerSex + ",");
                                else
                                {
                                    // Append default value
                                    sbClientList.Append("9,");
                                    // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                    if (String.IsNullOrEmpty(dbCarerSex) == false)
                                        sbErrorList.AppendLine("Invalid Carer Gender Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbCarerSex + "");
                                    else
                                        sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer Gender");
                                }

                                //Append country and language or advise of error
                                if (String.IsNullOrEmpty(dbCCountryCode) == false)
                                    sbClientList.Append(dbCCountryCode + ",");
                                else
                                {
                                    sbClientList.Append("9999,");
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer Country");
                                }
                                if (String.IsNullOrEmpty(dbCLanguageCode) == false)
                                    sbClientList.Append(dbCLanguageCode + ",");
                                else
                                {
                                    sbClientList.Append("9999,");
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer Language");
                                }

                                // Append carer indigenous status or advise of error
                                if (varCarerIndigenous > 0)
                                    sbClientList.Append(varCarerIndigenous + ",");
                                else
                                {
                                    // Append default value
                                    sbClientList.Append("99,");
                                    // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                    if (String.IsNullOrEmpty(dbCarerIndigenous) == false)
                                        sbErrorList.AppendLine("Invalid Carer Indigenous Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbCarerIndigenous + "");
                                    else
                                        sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer Indigenous State");
                                }

                                // Append carer state or advise of error
                                if (varCarerState > 0)
                                    sbClientList.Append(varCarerState + ",");
                                else
                                {
                                    // Append default value
                                    sbClientList.Append("9,");
                                    // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                    if (String.IsNullOrEmpty(dbCarerState) == false)
                                        sbErrorList.AppendLine("Invalid Carer State Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbCarerState + "");
                                    else
                                        sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer State");
                                }

                                // Append carer locality or advise of error
                                if (String.IsNullOrEmpty(dbCarerLocality) == false)
                                {
                                    // Locality must be less or equal to 46 characters
                                    if (dbCarerLocality.Length <= 46)
                                        sbClientList.Append(dbCarerLocality + ",");
                                    else
                                    {
                                        sbClientList.Append(",");
                                        sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer Locality");
                                    }
                                }
                                else
                                {
                                    sbClientList.Append(",");
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer Locality");
                                }

                                // Append carer postcode if numeric or advise of error
                                if (Int32.TryParse(dbCarerPostcode, out varCheckInt))
                                    sbClientList.Append(dbCarerPostcode + ",");
                                else
                                {
                                    sbClientList.Append("9999,");
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer Postcode");
                                }

                                // Append carer residence or advise of error
                                if (varCarerResidence > 0)
                                    sbClientList.Append(varCarerResidence + ",");
                                else
                                {
                                    // Append default value
                                    sbClientList.Append("9,");
                                    // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                    if (String.IsNullOrEmpty(dbCarerResidence) == false)
                                        sbErrorList.AppendLine("Invalid Carer Residence Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbCarerResidence + "");
                                    else
                                        sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer Residence");
                                }

                                // Append carer relationship or advise of error
                                if (varCarerRship > 0)
                                    sbClientList.Append(varCarerRship + ",");
                                else
                                {
                                    // Append default value
                                    sbClientList.Append("99,");
                                    // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                    if (String.IsNullOrEmpty(dbCarerRship) == false)
                                        sbErrorList.AppendLine("Invalid Carer Relationship Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbCarerRship + "");
                                    else
                                        sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer Relationship");
                                }

                                // Append carer for more or advise of error
                                if (varCarerMore > 0)
                                    sbClientList.Append(varCarerMore + ",");
                                else
                                {
                                    sbClientList.Append("99,");
                                    // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                    if (String.IsNullOrEmpty(dbCarerMore) == false)
                                        sbErrorList.AppendLine("Invalid Carer More People Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbCarerMore + "");
                                    else
                                        sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Carer More People");
                                }
                            }
                            else
                            {
                                //Leave blank if no carer
                                sbClientList.Append(",");
                                sbClientList.Append(",");
                                sbClientList.Append(",");
                                sbClientList.Append(",");
                                sbClientList.Append(",");
                                sbClientList.Append(",");
                                sbClientList.Append(",");
                                sbClientList.Append(",");
                                sbClientList.Append(",");
                                sbClientList.Append(",");
                                sbClientList.Append("0,");
                                sbClientList.Append("0,");
                                sbClientList.Append(",");
                            }

                            //Append date of last update info or advise of error
                            if (String.IsNullOrEmpty(dbModifiedOn) == false)
                            {
                                // Date must be greater than 1895
                                if (Convert.ToDateTime(dbModifiedOn).Year > 1895)
                                    sbClientList.Append(cleanDateM(Convert.ToDateTime(dbModifiedOn)) + ",");
                                else
                                {
                                    sbClientList.Append(",");
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Date of Last Update");
                                }

                            }
                            else
                            {
                                // Alert of error if missing
                                sbClientList.Append(",");
                                sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Date of Last Update");
                            }

                            // Append referral source or advise of error
                            if (varReferral > 0)
                                sbClientList.Append(varReferral + ",");
                            else
                            {
                                // Append default value
                                sbClientList.Append("99,");
                                // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                if (String.IsNullOrEmpty(dbSourceRef) == false)
                                    sbErrorList.AppendLine("Invalid Referral Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbSourceRef + "");
                                else
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Source of Referral");
                            }

                            //Append entry date information or advise of error
                            if (String.IsNullOrEmpty(dbStartDate) == false)
                            {
                                // Date must be greater than 1895
                                if (Convert.ToDateTime(dbStartDate).Year > 1895)
                                {
                                    sbClientList.Append(cleanDateM(Convert.ToDateTime(dbStartDate)) + ",");
                                }
                                else
                                {
                                    sbClientList.Append(",");
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Start Date");
                                }
                            }
                            else
                            {
                                // Alert of missing data
                                sbClientList.Append(",");
                                sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Start Date");
                            }


                            //Only append end date if it elapsed in the quarter
                            if (String.IsNullOrEmpty(dbEndDate) == false)
                            {
                                // If end date hasnt elapsed then append nothing, otherwise append end date if valid
                                if (Convert.ToDateTime(dbEndDate) > varEndDate)
                                {
                                    dbEndDate = "";
                                    sbClientList.Append(",");
                                }
                                else if (Convert.ToDateTime(dbEndDate).Year > 1895)
                                {
                                    sbClientList.Append(cleanDateM(Convert.ToDateTime(dbEndDate)) + ",");
                                }
                                else
                                {
                                    sbClientList.Append(",");
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Exit Date");
                                }

                            }
                            else
                            {
                                // Append nothing if empty
                                sbClientList.Append(",");
                            }

                            //Only append cessation information if there is an exit date
                            if (String.IsNullOrEmpty(dbEndDate) == false)
                            {
                                // Append cessation reason or advise of error
                                if (varCessation > 0)
                                    sbClientList.Append(varCessation + ",");
                                else
                                {
                                    // Append default value
                                    sbClientList.Append("99,");
                                    // If there is data then the value didnt match a drop down list value, otherwise advise of no data
                                    if (String.IsNullOrEmpty(dbCessation) == false)
                                        sbErrorList.AppendLine("Invalid Cessation Value: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - " + dbCessation + "");
                                    else
                                        sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Cessation Reason");
                                }

                            }
                            else
                                sbClientList.Append(",");

                            // This is not applicable to HHS and has been removed
                            // Ideally this should be added for completeness
                            //DomAssist
                            sbClientList.Append(",");
                            //VolSocial
                            sbClientList.Append(",");
                            //NursingH
                            sbClientList.Append(",");
                            //NursingC
                            sbClientList.Append(",");
                            //PodiatryH
                            sbClientList.Append(",");
                            //OccupationH
                            sbClientList.Append(",");
                            //SpeechH
                            sbClientList.Append(",");
                            //DieteticsH
                            sbClientList.Append(",");
                            //PhysioH
                            sbClientList.Append(",");
                            //AudiologyH
                            sbClientList.Append(",");
                            //CounsellingH
                            sbClientList.Append(",");
                            //AlliedH
                            sbClientList.Append(",");
                            //PodiatryC
                            sbClientList.Append(",");
                            //OccupationC
                            sbClientList.Append(",");
                            //SpeechC
                            sbClientList.Append(",");
                            //DieteticsC
                            sbClientList.Append(",");
                            //PhysioC
                            sbClientList.Append(",");
                            //AudiologyC
                            sbClientList.Append(",");
                            //CounsellingC
                            sbClientList.Append(",");
                            //AlliedC
                            sbClientList.Append(",");
                            //Personal
                            sbClientList.Append(",");
                            //PlannedCore
                            sbClientList.Append(",");
                            //PlannedHigh
                            sbClientList.Append(",");
                            //MealsH
                            sbClientList.Append(",");
                            //MealsC
                            sbClientList.Append(",");
                            //Respite
                            sbClientList.Append(",");
                            //Assessment
                            sbClientList.Append(",");
                            //CareMgt
                            sbClientList.Append(",");
                            //ClientCase
                            sbClientList.Append(",");
                            //PropMaint
                            sbClientList.Append(",");
                            //SelfCareAids
                            sbClientList.Append(",");
                            //MobilityAids
                            sbClientList.Append(",");
                            //CommAids
                            sbClientList.Append(",");
                            //ReadingAids
                            sbClientList.Append(",");
                            //MedicalAids
                            sbClientList.Append(",");
                            //CarMod
                            sbClientList.Append(",");
                            //GoodEquip
                            sbClientList.Append(",");
                            //CounselCare
                            sbClientList.Append(",");
                            //CounselCarer
                            sbClientList.Append(",");

                            //arrHousework = arrResultSet(80,iCounter)
                            sbClientList.Append("9,");
                            //arrTransport = arrResultSet(81,iCounter)
                            sbClientList.Append("9,");
                            //arrShopping = arrResultSet(82,iCounter)
                            sbClientList.Append("9,");
                            //arrMedication = arrResultSet(83,iCounter)
                            sbClientList.Append("9,");
                            //arrMoney = arrResultSet(84,iCounter)
                            sbClientList.Append("9,");
                            //arrWalking = arrResultSet(85,iCounter)
                            sbClientList.Append("9,");
                            //arrMobility = arrResultSet(86,iCounter)
                            sbClientList.Append("9,");

                            //arrSelfCare = arrResultSet(87,iCounter)
                            sbClientList.Append("2,");

                            //arrBathing = arrResultSet(88,iCounter)
                            sbClientList.Append("9,");
                            //arrDressing = arrResultSet(89,iCounter)
                            sbClientList.Append("9,");
                            //arrEating = arrResultSet(90,iCounter)
                            sbClientList.Append("9,");
                            //arrToilet = arrResultSet(91,iCounter)
                            sbClientList.Append("9,");
                            //arrCommunication = arrResultSet(92,iCounter)
                            sbClientList.Append("9,");

                            //arrMemory = arrResultSet(93,iCounter)
                            sbClientList.Append("2,");
                            //arrBehaviour = arrResultSet(94,iCounter)
                            sbClientList.Append("2,");

                            //arrHRSClient = arrResultSet(38,iCounter)
                            sbClientList.Append("0,");
                            //arrHRSConfirmation = arrResultSet(39,iCounter)
                            sbClientList.Append("9,");

                            //Time1
                            sbClientList.Append(",");
                            //Time2
                            sbClientList.Append(",");
                            //Time3
                            sbClientList.Append(",");
                            //Time4
                            sbClientList.Append(",");

                            //Reset the counter and activity data
                            varCheckTotal = 0;
                            varSCPDay = 0;
                            varSCPNightNon = 0;
                            varSCPNightActive = 0;
                            varSCPResidential = 0;
                            varSCPCounselSupport = 0;

                            varCCPAssertOut = 0;
                            varCCPCareCoord = 0;
                            varCCPHouseAss = 0;
                            varCCPGroupSocial = 0;

                            varHSAPAssertOut = 0;
                            varHSAPCareCoord = 0;
                            varHSAPHouseAss = 0;

                            varOPHRAssertOut = 0;
                            varOPHRCareCoord = 0;
                            varOPHRHouseAss = 0;
                            varOPHRGroupSocial = 0;

                            varSRSCareCoord = 0;
                            varSRSHouseAss = 0;
                            varSRSGroupSocial = 0;

                            // Loop through activities
                            foreach (var a in result7.Entities)
                            {
                                // We need to get the entity id for the client field for comparisons
                                if (a.Attributes.Contains("new_palmclientsupport1.new_client"))
                                {
                                    // Get the entity id for the client using the entity reference object
                                    getEntity = (EntityReference)a.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_client").Value;
                                    varClient2 = getEntity.Id.ToString();
                                }
                                else if (a.FormattedValues.Contains("new_palmclientsupport1.new_client"))
                                    varClient2 = a.FormattedValues["new_palmclientsupport1.new_client"];
                                else
                                    varClient2 = ".";

                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (a.FormattedValues.Contains("new_amount"))
                                    varAmount = a.FormattedValues["new_amount"];
                                else if (a.Attributes.Contains("new_amount"))
                                    varAmount = a.Attributes["new_amount"].ToString();
                                else
                                    varAmount = "";

                                // Convert amount to double or set to 0
                                Double.TryParse(varAmount, out arrAmount);

                                if (a.FormattedValues.Contains("new_activitymds"))
                                    varActivityMDS = a.FormattedValues["new_activitymds"];
                                else if (a.Attributes.Contains("new_activitymds"))
                                    varActivityMDS = a.Attributes["new_activitymds"].ToString();
                                else
                                    varActivityMDS = "";

                                if (a.FormattedValues.Contains("new_entrydate"))
                                    varEntryDate = a.FormattedValues["new_entrydate"];
                                else if (a.Attributes.Contains("new_entrydate"))
                                    varEntryDate = a.Attributes["new_entrydate"].ToString();
                                else
                                    varEntryDate = "";

                                // Convert American date format to Australian date format
                                varEntryDate = cleanDateAM(varEntryDate);

                                //foreach (KeyValuePair<String, Object> attribute in a.Attributes)
                                //{
                                //    varTest += (attribute.Key + ": " + attribute.Value + "\r\n");
                                //}
                                
                                // Only add data if the client ids match
                                if (dbPalmClientId == varClient2)
                                {
                                    // Add the amount to the specific total and overall total based on the activity type
                                    if (varActivityMDS == "SCP Respite Daytime in Home")
                                    {
                                        varSCPDay += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "SCP Respite Overnight in Home Non-active")
                                    {
                                        varSCPNightNon += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "SCP Respite Overnight in Home Active")
                                    {
                                        varSCPNightActive += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "SCP Residential Respite")
                                    {
                                        varSCPResidential += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "SCP Counselling & Support")
                                    {
                                        varSCPCounselSupport += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "CCP Assertive Outreach")
                                    {
                                        varCCPAssertOut += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "CCP Care Coordination")
                                    {
                                        varCCPCareCoord += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "CCP Housing Assistance")
                                    {
                                        varCCPHouseAss += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "CCP Group Social Support")
                                    {
                                        varCCPGroupSocial += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "HSAP Assertive Outreach")
                                    {
                                        varHSAPAssertOut += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "HSAP Care Coordination")
                                    {
                                        varHSAPCareCoord += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "HSAP Housing Assistance")
                                    {
                                        varHSAPHouseAss += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "OP High Rise Assertive Outreach")
                                    {
                                        varOPHRAssertOut += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "OP High Rise Care Coordination")
                                    {
                                        varOPHRCareCoord += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "OP High Rise Housing Assistance")
                                    {
                                        varOPHRHouseAss += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "OP High Rise Group Social Support")
                                    {
                                        varOPHRGroupSocial += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "SRS SC & Support Care Coordination")
                                    {
                                        varSRSCareCoord += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "SRS SC & Support Housing Assistance")
                                    {
                                        varSRSHouseAss += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }

                                    if (varActivityMDS == "SRS SC & Support Group Social Support")
                                    {
                                        varSRSGroupSocial += arrAmount;
                                        varCheckTotal += arrAmount;
                                    }
                                }
                            }

                            //If there are no activity hours, alert of error
                            if (varCheckTotal == 0)
                                sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - no activities");

                            //Reset financial data
                            varSCPGoodsEquip = 0;
                            varCCPFlexCare = 0;
                            varHSAPFlexCare = 0;
                            varOPHRFlexCare = 0;

                            // Loop through financial data
                            foreach (var f in result8.Entities)
                            {
                                // We need to get the entity id for the client field for comparisons
                                if (f.Attributes.Contains("new_palmclientsupport1.new_client"))
                                {
                                    // Get the entity id for the client using the entity reference object
                                    getEntity = (EntityReference)f.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_client").Value;
                                    varClient2 = getEntity.Id.ToString();
                                }
                                else if (f.FormattedValues.Contains("new_palmclientsupport1.new_client"))
                                    varClient2 = f.FormattedValues["new_palmclientsupport1.new_client"];
                                else
                                    varClient2 = ".";

                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (f.FormattedValues.Contains("new_amount"))
                                    varAmount = f.FormattedValues["new_amount"];
                                else if (f.Attributes.Contains("new_amount"))
                                    varAmount = f.Attributes["new_amount"].ToString();
                                else
                                    varAmount = "";

                                // Convert amount to double or set to 0
                                Double.TryParse(varAmount, out arrAmount);

                                if (f.FormattedValues.Contains("new_mds"))
                                    varMDS = f.FormattedValues["new_mds"];
                                else if (f.Attributes.Contains("new_mds"))
                                    varMDS = f.Attributes["new_mds"].ToString();
                                else
                                    varMDS = "";

                                if (f.FormattedValues.Contains("new_entrydate"))
                                    varEntryDate = f.FormattedValues["new_entrydate"];
                                else if (f.Attributes.Contains("new_entrydate"))
                                    varEntryDate = f.Attributes["new_entrydate"].ToString();
                                else
                                    varEntryDate = "";

                                // Convert American Date format to Australian Date format
                                varEntryDate = cleanDateAM(varEntryDate);

                                // Only add financial if the client numbers are the same
                                if (dbPalmClientId == varClient2)
                                {
                                    //(Comment out SCP as not relevant)
                                    //if (varMDS == "Support for Carers Goods & Equipment")
                                    //varSCPGoodsEquip += arrAmount;

                                    // Add the amount to the specific total and overall total based on the financial type
                                    if (varMDS == "CCP Flexible Care Funds")
                                        varCCPFlexCare += arrAmount;

                                    if (varMDS == "HSAP Flexible Care Funds")
                                        varHSAPFlexCare += arrAmount;

                                    if (varMDS == "OP High Rise Flexible Care Funds")
                                        varOPHRFlexCare += arrAmount;
                                }
                            }

                            //Round up the values for each amount using round up function below
                            varSCPDay = roundUp(varSCPDay);
                            varSCPNightNon = roundUp(varSCPNightNon);
                            varSCPNightActive = roundUp(varSCPNightActive);
                            varSCPResidential = roundUp(varSCPResidential);
                            varSCPCounselSupport = roundUp(varSCPCounselSupport);
                            varSCPGoodsEquip = roundUp(varSCPGoodsEquip);
                            varCCPAssertOut = roundUp(varCCPAssertOut);
                            varCCPCareCoord = roundUp(varCCPCareCoord);
                            varCCPFlexCare = roundUp(varCCPFlexCare);
                            varCCPHouseAss = roundUp(varCCPHouseAss);
                            varCCPGroupSocial = roundUp(varCCPGroupSocial);
                            varHSAPAssertOut = roundUp(varHSAPAssertOut);
                            varHSAPCareCoord = roundUp(varHSAPCareCoord);
                            varHSAPFlexCare = roundUp(varHSAPFlexCare);
                            varHSAPHouseAss = roundUp(varHSAPHouseAss);
                            varOPHRAssertOut = roundUp(varOPHRAssertOut);
                            varOPHRCareCoord = roundUp(varOPHRCareCoord);
                            varOPHRFlexCare = roundUp(varOPHRFlexCare);
                            varOPHRHouseAss = roundUp(varOPHRHouseAss);
                            varOPHRGroupSocial = roundUp(varOPHRGroupSocial);
                            varSRSCareCoord = roundUp(varSRSCareCoord);
                            varSRSHouseAss = roundUp(varSRSHouseAss);
                            varSRSGroupSocial = roundUp(varSRSGroupSocial);

                            //Append the hours and dollars for each type
                            sbClientList.Append(varSCPDay + ",");
                            sbClientList.Append(varSCPNightNon + ",");
                            sbClientList.Append(varSCPNightActive + ",");
                            sbClientList.Append(varSCPResidential + ",");
                            sbClientList.Append(varSCPCounselSupport + ",");
                            sbClientList.Append(varSCPGoodsEquip + ",");

                            sbClientList.Append(varCCPDepChild + ",");
                            sbClientList.Append(varCCPDisType + ",");

                            sbClientList.Append(varCCPAssertOut + ",");
                            sbClientList.Append(varCCPCareCoord + ",");
                            sbClientList.Append(varCCPFlexCare + ",");
                            sbClientList.Append(varCCPHouseAss + ",");
                            sbClientList.Append(varCCPGroupSocial + ",");
                            sbClientList.Append(varHSAPAssertOut + ",");
                            sbClientList.Append(varHSAPCareCoord + ",");
                            sbClientList.Append(varHSAPFlexCare + ",");
                            sbClientList.Append(varHSAPHouseAss + ",");
                            sbClientList.Append(varOPHRAssertOut + ",");
                            sbClientList.Append(varOPHRCareCoord + ",");
                            sbClientList.Append(varOPHRFlexCare + ",");
                            sbClientList.Append(varOPHRHouseAss + ",");
                            sbClientList.Append(varOPHRGroupSocial + ",");
                            sbClientList.Append(varSRSCareCoord + ",");
                            sbClientList.Append(varSRSHouseAss + ",");
                            sbClientList.Append(varSRSGroupSocial + ",");

                            // Close the client row
                            sbClientList.AppendLine("ENDCLIENT");
                        }

                        // Create Header part of the MDS extract
                        sbHeaderList.Append("HACC,");
                        sbHeaderList.Append("201,");
                        sbHeaderList.Append("HEADER,");
                        sbHeaderList.Append("3580,");
                        sbHeaderList.Append(varYear + "/" + varQuarter + ",");
                        sbHeaderList.Append("1,");
                        sbHeaderList.Append("1,");
                        sbHeaderList.Append(varClientCount + ",");
                        sbHeaderList.Append("LMHS,");
                        sbHeaderList.AppendLine("ENDHEADER");
                        // Add Client data
                        sbHeaderList.AppendLine(sbClientList.ToString());
                        
                        // Create note against current Palm Go MDS record and add attachment
                        byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                        string encodedData = System.Convert.ToBase64String(filename);
                        Entity Annotation = new Entity("annotation");
                        Annotation.Attributes["objectid"] = new EntityReference("new_palmgomds", varMdsID);
                        Annotation.Attributes["objecttypecode"] = "new_palmgomds";
                        Annotation.Attributes["subject"] = "MDS Extract";
                        Annotation.Attributes["documentbody"] = encodedData;
                        Annotation.Attributes["mimetype"] = @"text/csv";
                        Annotation.Attributes["notetext"] = "MDS Extract for " + varYear + " quarter " + varQuarter;
                        Annotation.Attributes["filename"] = varFileName;
                        _service.Create(Annotation);

                        // If there is an error, create note against current Palm Go MDS record and add attachment
                        if (sbErrorList.Length > 0)
                        {
                            byte[] filename2 = Encoding.ASCII.GetBytes(sbErrorList.ToString());
                            string encodedData2 = System.Convert.ToBase64String(filename2);
                            Entity Annotation2 = new Entity("annotation");
                            Annotation2.Attributes["objectid"] = new EntityReference("new_palmgomds", varMdsID);
                            Annotation2.Attributes["objecttypecode"] = "new_palmgomds";
                            Annotation2.Attributes["subject"] = "MDS Extract";
                            Annotation2.Attributes["documentbody"] = encodedData2;
                            Annotation2.Attributes["mimetype"] = @"text/csv";
                            Annotation2.Attributes["notetext"] = "MDS errors and warnings for " + varYear + " quarter " + varQuarter;
                            Annotation2.Attributes["filename"] = varFileName2;
                            _service.Create(Annotation2);
                        }

                        // Debug
                        // throw new InvalidPluginExecutionException("This plugin is working:\r\n" + varTest);
                    }


                }
                // Error if plugin code fails
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in the FollowupPlugin plug-in.", ex);
                }
                // Error if plugin code fails
                catch (Exception ex)
                {
                    tracingService.Trace("FollowupPlugin: {0}", ex.ToString());
                    throw;
                }
            }
        }

        //Limit a string to characters
        public string cleanString(string clean, string thetype)
        {
            string varCharAllowed = ""; //Characters allower
            string temp = ""; //Temporary string for removing illegal characters

            //Set the characters allowed in the string
            if (thetype == "normal")
                varCharAllowed = "abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ-_1234567890(),"; //Characters allowed
            else if (thetype == "number")
                varCharAllowed = "1234567890"; //Characters allowed
            else if (thetype == "phone")
                varCharAllowed = "1234567890 ()"; //Characters allowed
            else if (thetype == "address")
                varCharAllowed = "abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ-_1234567890(),/"; //Characters allowed
            else if (thetype == "double")
                varCharAllowed = "1234567890."; //Characters allowed
            else if (thetype == "numstring")
                varCharAllowed = "1234567890,"; //Characters allowed
            else if (thetype == "name")
                varCharAllowed = "abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ-'";
            else if (thetype == "username")
                varCharAllowed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            else if (thetype == "mailbox")
                varCharAllowed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890-.";
            else if (thetype == "slk")
                varCharAllowed = "ABCDEFGHIJKLMNOPQRSTUVWXYZ29";
            else if (thetype == "letter")
                varCharAllowed = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
            else if (thetype == "search")
                varCharAllowed = "abcdefghijklmnopqrstuvwxyz1234567890.";
            else if (thetype == "palm")
                varCharAllowed = "PLMplm1234567890";
            else if (thetype == "voucher")
                varCharAllowed = "abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ-_()1234567890#";
            else
                varCharAllowed = "abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ-_1234567890()"; //Characters allowed

            //Set a temporary string to the value of the string passed
            temp = clean;

            if (String.IsNullOrEmpty(clean) == false)
            {
                clean = clean.Trim();
                temp = clean;
            }

            if (String.IsNullOrEmpty(clean) == false)
            {

                //Loop through each character in the string
                for (int i = 0; i < clean.Length; i++)
                {
                    //If the next character is not in allowed set of characters, replace it with ~
                    if (varCharAllowed.IndexOf(clean[i]) == -1 && clean[i].ToString() != "~")
                        temp = temp.Replace(clean[i].ToString(), "~");
                }

                //Set the string to the value of the string, minus the cleaned characters
                clean = temp.Replace("~", "");

            }

            return clean;
        }

        // Format: 01/01/1970
        public string cleanDateM(DateTime getdate)
        {
            string varDay = "";
            string varMonth = "";

            if (getdate.Day < 10)
                varDay = "0" + getdate.Day;
            else
                varDay = getdate.Day + "";

            if (getdate.Month < 10)
                varMonth = "0" + getdate.Month;
            else
                varMonth = getdate.Month + "";

            string clean = varDay + "/" + varMonth + "/" + getdate.Year;
            return clean;
        }

        // Format: 1-Jan-1970
        public string cleanDate(DateTime getdate)
        {
            string clean = getdate.Day + "-" + getdate.ToString("MMM") + "-" + getdate.Year;
            return clean;
        }

        // Fix American Date issue
        public string cleanDateAM(string getdate)
        {
            string[] test;
            int count = 0;
            string varD = "";
            string varM = "";
            string varY = "";
            string clean = "";
            DateTime varCheckDate = new DateTime();

            if (getdate.IndexOf(" ") > 0)
                getdate = getdate.Substring(0, getdate.IndexOf(" "));

            test = getdate.Split('/');

            foreach (string dt in test)
            {
                if (count == 0)
                    varD = dt;
                if (count == 1)
                    varM = dt;
                if (count == 2)
                    varY = dt;

                count++;
            }

            if (String.IsNullOrEmpty(varD) == false && String.IsNullOrEmpty(varM) == false && String.IsNullOrEmpty(varY) == false)
            {
                if (varD.Substring(0, 1) == "0")
                    varD = varD.Replace("0", "");
                if (varM.Substring(0, 1) == "0")
                    varM = varM.Replace("0", "");

                if (varM == "1")
                    varM = "Jan";
                else if (varM == "2")
                    varM = "Feb";
                else if (varM == "3")
                    varM = "Mar";
                else if (varM == "4")
                    varM = "Apr";
                else if (varM == "5")
                    varM = "May";
                else if (varM == "6")
                    varM = "Jun";
                else if (varM == "7")
                    varM = "Jul";
                else if (varM == "8")
                    varM = "Aug";
                else if (varM == "9")
                    varM = "Sep";
                else if (varM == "10")
                    varM = "Oct";
                else if (varM == "11")
                    varM = "Nov";
                else if (varM == "12")
                    varM = "Dec";

                clean = varD + "-" + varM + "-" + varY;
            }

            if (!DateTime.TryParse(clean, out varCheckDate))
                clean = "";

            return clean;
        }

        // Round up number
        public double roundUp(double value)
        {
            string checkValue = value + "";

            if (checkValue.IndexOf(".") > -1)
            {
                checkValue = checkValue.Substring(0, checkValue.IndexOf("."));

                value = Convert.ToDouble(checkValue);
                value++;
            }

            return value;
        }

        // Convert multiselect option set to values with asterisks around them for better string matching
        public string getMult(string s)
        {
            string[] myMult;
            string getVal = "";

            if (String.IsNullOrEmpty(s) == false)
            {
                if (s.IndexOf(";") > -1)
                {
                    myMult = s.Split(';');
                    foreach (string tst in myMult)
                    {
                        getVal += "*" + tst.Trim() + "*,";
                    }
                }
                else
                    getVal = "*" + s + "*";
            }

            return getVal;
        }
    }
}


