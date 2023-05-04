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
	public class goDEX: IPlugin
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

                // Verify that the target entity represents the Palm Go DEX entity
                // If not, this plug-in was not registered correctly.
                // if (entity.LogicalName != "new_palmgodex")
                //    return;

                try
                {
                    // Fixes:
                    // Remove ampersand from comments
                    // Compare values to Lower

                    string varDescription = ""; // Description field on form
                    bool varCreateExtract = false; // Create extract field on form
                    string varAgencyId = ""; // Agency Id field on form
                    string varOutletActivityId = ""; // Agency name
                    string varReport = ""; //Report Type
                    Guid varDexID = new Guid(); // GUID for palm go dex record
                    StringBuilder sbHeaderList = new StringBuilder(); // Header part of extract
                    StringBuilder sbClientList = new StringBuilder(); // Client List part of extract
                    StringBuilder sbErrorList = new StringBuilder(); // Error list
                    StringBuilder sbCaseList = new StringBuilder(); // Case part of extract
                    StringBuilder sbSessionList = new StringBuilder(); // Session part of extract
                    StringBuilder sbSessionAssessmentList = new StringBuilder(); // Session Assessments part of extract
                    StringBuilder sbSessionAssessment2List = new StringBuilder(); // Sub Session Assessments part of extract
                    StringBuilder sbClientAssessmentList = new StringBuilder(); // Client Assessments part of extract
                    StringBuilder sbClientAssessment2List = new StringBuilder(); // Sub Client Assessments part of extract
                    string varFileName = ""; // Extract file name
                    string varFileName2 = ""; // Error log file name
                    DateTime varStartDate = new DateTime(); // Start date of extract
                    DateTime varEndDate = new DateTime(); // End date of extract
                    DateTime varStartDatePr = new DateTime(); // Printable Start date of extract
                    DateTime varEndDatePr = new DateTime(); // Printable End date of extract

                    int varCheckInt = 0; // Used to parse integers
                    double varCheckDouble = 0; // Used to parse doubles
                    DateTime varCheckDate = new DateTime(); // Used to parse dates
                    string varExt = ""; // Extension for file and ids
                    string varDEXType = ""; // Type of DEX record
                    bool varDoCHSP = false; // Whether the DEX is for CHSP
                    EntityReference getEntity; // Entity reference object
                    AliasedValue getAlias; // Aliased value object

                    string varTest = ""; // Used for debug

                    // Only do this if the entity is the Palm Go DEX entity
                    if (entity.LogicalName == "new_palmgodex")
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

                        // Get the report type
                        if (entity.Contains("new_report"))
                            varReport = entity.FormattedValues["new_report"];

                        // Get info from current Palm Go DEX record
                        varDescription = entity.GetAttributeValue<string>("new_description");
                        varCreateExtract = entity.GetAttributeValue<bool>("new_createextract");

                        // Important: The plugin uses American dates but returns formatted Australian dates
                        // Any dates created in the plugin will be American

                        varStartDate = entity.GetAttributeValue<DateTime>("new_datefrom");
                        varStartDatePr = Convert.ToDateTime(varStartDate.AddHours(14).ToString()); // Australian Date

                        varEndDate = entity.GetAttributeValue<DateTime>("new_dateto");
                        varEndDatePr = Convert.ToDateTime(varEndDate.AddHours(23).ToString()); // Australian Date

                        varEndDate = varEndDate.AddHours(23); // Correct for Australian time

                        //varTest += varStartDate + " " + varEndDate;

                        varDexID = entity.Id; // Get GUID

                        // Get associated details from DEX agency
                        EntityReference ownerLookup = (EntityReference)entity.Attributes["new_agencyid"];
                        varAgencyId += ownerLookup.Id.ToString() + ".\r\n";
                        varAgencyId += ((EntityReference)entity.Attributes["new_agencyid"]).Name + ".\r\n";
                        varAgencyId += ownerLookup.LogicalName + ".\r\n";

                        var actualOwningUnit = _service.Retrieve(ownerLookup.LogicalName, ownerLookup.Id, new ColumnSet(true));

                        // Get agency name and extension
                        varAgencyId = actualOwningUnit["new_agency"].ToString();
                        varExt = actualOwningUnit["new_ext"].ToString();

                        // Get DEX type
                        if (actualOwningUnit.FormattedValues.Contains("new_dextype"))
                            varDEXType = actualOwningUnit.FormattedValues["new_dextype"];

                        // Get outlet activity id
                        varOutletActivityId = actualOwningUnit["new_outletactivityid"].ToString();

                        // Determine if agency is for CHSP
                        if (varDEXType.ToLower() == "chsp")
                            varDoCHSP = true;

                        //varTest += varAgencyId + " - " + varExt + " - " + varDEXType + " - " + varDoCHSP;

                        // Create file name
                        varFileName = "DEX_" + varStartDatePr.Year;

                        if (varStartDatePr.Month < 10)
                            varFileName += "0" + varStartDatePr.Month;
                        else
                            varFileName += varStartDatePr.Month + "";

                        if (varStartDatePr.Day < 10)
                            varFileName += "0" + varStartDatePr.Day;
                        else
                            varFileName += varStartDatePr.Day;

                        varFileName += "_" + varEndDatePr.Year;

                        if (varEndDatePr.Month < 10)
                            varFileName += "0" + varEndDatePr.Month;
                        else
                            varFileName += varEndDatePr.Month + "";

                        if (varEndDatePr.Day < 10)
                            varFileName += "0" + varEndDatePr.Day;
                        else
                            varFileName += varEndDatePr.Day;

                        varFileName += "_" + varExt + ".xml";

                        // Create file name for errors
                        varFileName2 = varFileName.Replace(".xml", ".txt");
                        varFileName2 = "Errors for " + varFileName2;

                        if (varReport == ("Internal"))
                        {
                            // Fetch statements for database
                            // Get the required fields from the client table (and associated entities)
                            // Any clients that have a DEX record for the period
                            string dbClientList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclient'>
                                <attribute name='new_palmclientid' />
                                <attribute name='new_address' />
                                <attribute name='new_firstname' />
                                <attribute name='new_surname' />
                                <attribute name='new_dob' />
                                <attribute name='new_dobest' />
                                <attribute name='new_gender' />
                                <attribute name='new_dexslk' />
                                <attribute name='new_country' />
                                <attribute name='new_language' />
                                <attribute name='new_indigenous' />
                                <attribute name='new_yeararrival' />
                                <link-entity name='new_palmclientsupport' to='new_palmclientid' from='new_client' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_ccpdistype' />
                                    <attribute name='new_locality' />
                                    <attribute name='new_ishomeless' />
                                    <attribute name='new_livingarrangepres' />
                                    <attribute name='new_incomepres' />
                                    <attribute name='new_interpreterused' />
                                    <attribute name='new_sourceref' />
                                    <attribute name='new_primreason' />
                                    <attribute name='new_reasons' />
                                    <attribute name='new_residentialpres' />
                                    <attribute name='new_dva' />
                                    <attribute name='new_careravail' />
                                    <attribute name='new_startdate' />
                                    <order attribute='new_startdate' descending='true' />
                                    <link-entity name='new_palmclientdex' to='new_palmclientsupportid' from='new_supportperiod' link-type='inner'>
                                        <attribute name='new_entrydate' />
                                        <attribute name='new_agencyid' />
                                    </link-entity>
                                    <link-entity name='new_palmddllocality' to='new_locality' from='new_palmddllocalityid' link-type='outer'>
                                        <attribute name='new_postcode' />
                                        <attribute name='new_state' />
                                    </link-entity>
                                </link-entity>
                                <link-entity name='new_palmddlcountry' to='new_country' from='new_palmddlcountryid' link-type='outer'>
                                    <attribute name='new_code' />
                                </link-entity>
                                <link-entity name='new_palmddllanguage' to='new_language' from='new_palmddllanguageid' link-type='outer'>
                                    <attribute name='new_code' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname = 'new_palmclientdex' attribute = 'new_agencyid' operator= 'eq' value = '" + ownerLookup.Id + @"' />
                                </filter>
                              </entity>
                            </fetch> ";

                            // Get the required fields from the dex table
                            // Any DEX records against the agency id that fall within the period
                            string dbDEXList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientdex'>
                                <attribute name='new_entrydate' />
                                <attribute name='new_palmclientdexid' />
                                <attribute name='new_consentfuture' />
                                <attribute name='new_consentrelease' />
                                <attribute name='new_norecip' />
                                <attribute name='new_hours' />
                                <attribute name='new_servtype' />
                                <attribute name='new_supportperiod' />
                                <attribute name='new_agencyid' />
                                <filter type='and'>
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientdex' attribute='new_agencyid' operator='eq' value='" + ownerLookup.Id + @"' />
                                </filter>
                                <order attribute='new_supportperiod' />
                              </entity>
                            </fetch> ";

                            // Get the required fields from the dex scores table
                            // Any DEX scores records against the agency id that fall within the period
                            string dbScoresList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientdex'>
                                <attribute name='new_entrydate' />
                                <attribute name='new_palmclientdexid' />
                                <attribute name='new_consentfuture' />
                                <attribute name='new_consentrelease' />
                                <attribute name='new_norecip' />
                                <attribute name='new_hours' />
                                <attribute name='new_servtype' />
                                <attribute name='new_supportperiod' />
                                <attribute name='new_agencyid' />
                                <link-entity name='new_palmclientdexscores' to='new_palmclientdexid' from='new_dexrecord' link-type='inner'>
                                    <attribute name='new_palmclientdexscoresid' />
                                    <attribute name='new_score' />
                                    <attribute name='new_section' />
                                    <attribute name='new_dexrecord' />
                                    <attribute name='new_thetype' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientdex' attribute='new_agencyid' operator='eq' value='" + ownerLookup.Id + @"' />
                                </filter>
                                <order attribute='new_supportperiod' />
                              </entity>
                            </fetch> ";

                            // Get the required fields from the SU Drop Down list entity
                            string dbDropDownList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                  <entity name='new_palmsudropdown'>
                                    <attribute name='new_type' />
                                    <attribute name='new_description' />
                                    <attribute name='new_er' />
                                    <order attribute='new_description' />
                                  </entity>
                                </fetch> ";

                            // Get the maximum number of recipients for each support period
                            string dbRecipList = @"
                           <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                              <entity name='new_palmclientdex'>
                                <attribute name='new_supportperiod' alias='new_supportperiod_max' groupby='true' />
                                <attribute name='new_norecip' alias='new_norecip_max' aggregate='max' />
                                <filter type='and'>
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='ge' value='1-Jul-2018' />
                                </filter>
                              </entity>
                            </fetch> ";

                            // Variables to hold database values
                            string dbPalmClientId = "";
                            string varPalmClientId = "";
                            string dbClient = "";
                            string dbFirstName = "";
                            string dbSurname = "";
                            string dbDob = "";
                            string dbDobEst = "";
                            string dbGender = "";
                            string dbDexSlk = "";
                            string dbCountry = "";
                            string dbLanguage = "";
                            string dbIndigenous = "";
                            string dbYearArrival = "";
                            string dbPalmClientSupportId = "";
                            string varPalmClientSupportId = "";
                            string dbCCPDisType = "";
                            string dbLocality = "";
                            string dbIsHomeless = "";
                            string dbLivingArrangePres = "";
                            string dbIncomePres = "";
                            string dbInterpreterUsed = "";
                            string dbSourceRef = "";
                            string dbPrimReason = "";
                            string dbReasons = "";
                            string dbResidentialPres = "";
                            string dbDva = "";
                            string dbCarerAvail = "";
                            string dbEntryDate = "";
                            string varEntryDate = "";
                            string varEntryDate2 = "";
                            string dbAgencyId = "";
                            string dbState = "";
                            string dbPostcode = "";

                            string dbEntryDate2 = "";
                            string dbPalmClientDexId = "";
                            string varPalmClientDexId = "";
                            string dbConsentFuture = "";
                            string dbConsentRelease = "";
                            string dbNoRecip = "";
                            string dbHours = "";
                            string dbServType = "";
                            string dbSupportPeriod = "";
                            string dbSupportPeriod2 = "";
                            string dbAgencyId2 = "";
                            string dbPalmClientDexScoresId = "";
                            string dbScore = "";
                            string dbSection = "";
                            string varSection = "";
                            string dbDexRecord = "";
                            string dbTheType = "";

                            string dbSupportPeriod_Max = "";
                            string dbNoRecip_Max = "";

                            // Drop down list variables
                            string varDesc = "";
                            string varType = "";
                            string varDEX = "";

                            // Variables for SLK
                            string varFirstName = "";
                            string varSurname = "";
                            string varGender = "";
                            string varDob = "";

                            // Variables for DEX record used to convert values
                            string varInterpreterUsed = "";
                            string varHasDis = "";
                            string varIsHomeless = "";

                            // Id variables
                            string varClientNumber = "";
                            string varSessionNumber = "";
                            string varCaseNumber = "";

                            // Used to determine whether to process records
                            int varDoNext = 0;
                            int varDoNext2 = 0;
                            int varDoNext3 = 0;
                            int varMaxRecip = 0; // Maximum recipients

                            string varReasonStr = ""; // Reasons for presenting string

                            // Validate whether data was found
                            bool varSource = false;
                            bool varIndigenous = false;
                            bool varState = false;
                            bool varIncome = false;
                            bool varLiving = false;
                            bool varCCP = false;
                            bool varResidential = false;
                            bool varDV = false;

                            int varNoRecip = 0; // Number of recipients
                            int varGetRecip = 0; // Recipients from table

                            // Strings for adding DEX scores
                            string varPreGroup = "a";
                            string varPostGroup = "a";
                            string varPreCirc = "a";
                            string varPostCirc = "a";
                            string varPreGoal = "a";
                            string varPostGoal = "a";
                            string varSatisfaction = "a";
                            bool varSeeClient3 = false; // If client exists
                            string dbPalmClientDexId2 = ""; // Dex ID for comparison
                            string[] value; // Values string
                            string varServTypeId = ""; // Service type id

                            //Debug variable
                            string dbTemp = "";

                            // Get the fetch XML data and place in entity collection objects
                            EntityCollection result = _service.RetrieveMultiple(new FetchExpression(dbClientList));
                            EntityCollection result2 = _service.RetrieveMultiple(new FetchExpression(dbDEXList));
                            EntityCollection result3 = _service.RetrieveMultiple(new FetchExpression(dbDropDownList));
                            EntityCollection result6 = _service.RetrieveMultiple(new FetchExpression(dbRecipList));
                            EntityCollection result8 = _service.RetrieveMultiple(new FetchExpression(dbScoresList));

                            // Strings to check for duplicates
                            varClientNumber = "...";
                            varSessionNumber = "...";
                            varCaseNumber = "...";

                            // Loop through client records
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

                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (c.FormattedValues.Contains("new_palmclientid"))
                                    dbPalmClientId = c.FormattedValues["new_palmclientid"];
                                else if (c.Attributes.Contains("new_palmclientid"))
                                    dbPalmClientId = c.Attributes["new_palmclientid"].ToString();
                                else
                                    dbPalmClientId = "";

                                // Remove dashes
                                varPalmClientId = "";
                                if (String.IsNullOrEmpty(dbPalmClientId) == false)
                                    varPalmClientId = dbPalmClientId.Replace("-", "");

                                if (c.FormattedValues.Contains("new_address"))
                                    dbClient = c.FormattedValues["new_address"];
                                else if (c.Attributes.Contains("new_address"))
                                    dbClient = c.Attributes["new_address"].ToString();
                                else
                                    dbClient = "";

                                //Reset duplicate client checker
                                varDoNext = 0;
                                

                                if (varClientNumber.IndexOf("*" + dbPalmClientId + "*") > -1)
                                {
                                    varDoNext = 1;
                                }


                                varClientNumber += "*" + dbPalmClientId + "*";

                                if (c.FormattedValues.Contains("new_firstname"))
                                    dbFirstName = c.FormattedValues["new_firstname"];
                                else if (c.Attributes.Contains("new_firstname"))
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

                                if (dbDobEst == "Yes" || dbDobEst == "true")
                                    dbDobEst = "true";
                                else
                                    dbDobEst = "false";

                                // Convert DOB to valid format
                                if (DateTime.TryParse(dbDob, out varCheckDate))
                                    dbDob = cleanDateE(varCheckDate);
                                else
                                {
                                    varCheckDate = Convert.ToDateTime("1-Jan-1970");
                                    dbDob = "1970-01-01";
                                }

                                if (c.FormattedValues.Contains("new_gender"))
                                    dbGender = c.FormattedValues["new_gender"];
                                else if (c.Attributes.Contains("new_gender"))
                                    dbGender = c.Attributes["new_gender"].ToString();
                                else
                                    dbGender = "";

                                if (c.FormattedValues.Contains("new_dexslk"))
                                    dbDexSlk = c.FormattedValues["new_dexslk"];
                                else if (c.Attributes.Contains("new_dexslk"))
                                    dbDexSlk = c.Attributes["new_dexslk"].ToString();
                                else
                                    dbDexSlk = "";



                                // Create SLK based on firstname, surname, gender and dob
                                varSurname = dbSurname.Replace(" ", string.Empty); //Prepare surname by removing spaces.
                                                                                   //Need to clean surname.
                                varSurname = cleanString(varSurname.ToUpper(), "slk");
                                varSurname = varSurname + "22222";
                                varSurname = varSurname.Substring(1, 1) + varSurname.Substring(2, 1) + varSurname.Substring(4, 1);

                                if (varSurname == "222")
                                    varSurname = "999";

                                varFirstName = dbFirstName.Replace(" ", string.Empty); //Prepare surname by removing spaces.
                                                                                       //Need to clean first name.
                                varFirstName = cleanString(varFirstName.ToUpper(), "slk");
                                varFirstName = varFirstName + "22222";
                                varFirstName = varFirstName.Substring(1, 1) + varFirstName.Substring(2, 1);

                                if (varFirstName == "22")
                                    varFirstName = "99";

                                //Get the gender code
                                if (dbGender == "Female")
                                    varGender = "2";
                                else if (dbGender == "Male")
                                    varGender = "1";
                                else
                                    varGender = "9";

                                //Put dob into expected format
                                varDob = cleanDateS(varCheckDate);

                                //Get the statistical linkage key
                                dbDexSlk = varSurname + varFirstName + varDob + varGender;

                                dbDexSlk = dbDexSlk.ToUpper();
                                if (dbDexSlk.Length != 14)
                                    sbErrorList.AppendLine("(" + dbClient + ") " + dbFirstName + " " + dbSurname + ": SLK not 14 characters long | " + dbDexSlk + "<br>");

                                if (c.FormattedValues.Contains("new_palmddlcountry4.new_code"))
                                    dbCountry = c.FormattedValues["new_palmddlcountry4.new_code"];
                                else if (c.Attributes.Contains("new_palmddlcountry4.new_code"))
                                    dbCountry = c.Attributes["new_palmddlcountry4.new_code"].ToString();
                                else
                                    dbCountry = "";

                                if (c.FormattedValues.Contains("new_palmddllanguage5.new_code"))
                                    dbLanguage = c.FormattedValues["new_palmddllanguage5.new_code"];
                                else if (c.Attributes.Contains("new_palmddllanguage5.new_code"))
                                    dbLanguage = c.Attributes["new_palmddllanguage5.new_code"].ToString();
                                else
                                    dbLanguage = "";

                                //varTest += dbCountry + " " + dbLanguage + "\r\n";

                                if (c.FormattedValues.Contains("new_indigenous"))
                                    dbIndigenous = c.FormattedValues["new_indigenous"];
                                else if (c.Attributes.Contains("new_indigenous"))
                                    dbIndigenous = c.Attributes["new_indigenous"].ToString();
                                else
                                    dbIndigenous = "";

                                if (c.FormattedValues.Contains("new_yeararrival"))
                                    dbYearArrival = c.FormattedValues["new_yeararrival"];
                                else if (c.Attributes.Contains("new_yeararrival"))
                                    dbYearArrival = c.Attributes["new_yeararrival"].ToString();
                                else
                                    dbYearArrival = "";

                                // Ensure numeric
                                dbYearArrival = cleanString(dbYearArrival, "number");

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_palmclientsupportid"))
                                    dbPalmClientSupportId = c.FormattedValues["new_palmclientsupport1.new_palmclientsupportid"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_palmclientsupportid"))
                                    dbPalmClientSupportId = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_palmclientsupportid").Value.ToString();
                                else
                                    dbPalmClientSupportId = "";

                                // Remove dashes
                                varPalmClientSupportId = "";
                                if (String.IsNullOrEmpty(dbPalmClientSupportId) == false)
                                    varPalmClientSupportId = dbPalmClientSupportId.Replace("-", "");

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_ccpdistype"))
                                    dbCCPDisType = c.FormattedValues["new_palmclientsupport1.new_ccpdistype"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_ccpdistype"))
                                    dbCCPDisType = c.Attributes["new_palmclientsupport1.new_ccpdistype"].ToString();
                                else
                                    dbCCPDisType = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_locality"))
                                    dbLocality = c.FormattedValues["new_palmclientsupport1.new_locality"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_locality"))
                                    dbLocality = c.Attributes["new_palmclientsupport1.new_locality"].ToString();
                                else
                                    dbLocality = "";

                                if (c.FormattedValues.Contains("new_palmddllocality3.new_state"))
                                    dbState = c.FormattedValues["new_palmddllocality3.new_state"];
                                else if (c.Attributes.Contains("new_palmddllocality3.new_state"))
                                    dbState = c.Attributes["new_palmddllocality3.new_state"].ToString();
                                else
                                    dbState = "";

                                if (c.FormattedValues.Contains("new_palmddllocality3.new_postcode"))
                                    dbPostcode = c.FormattedValues["new_palmddllocality3.new_postcode"];
                                else if (c.Attributes.Contains("new_palmddllocality3.new_postcode"))
                                    dbPostcode = c.Attributes["new_palmddllocality3.new_postcode"].ToString();
                                else
                                    dbPostcode = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_ishomeless"))
                                    dbIsHomeless = c.FormattedValues["new_palmclientsupport1.new_ishomeless"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_ishomeless"))
                                    dbIsHomeless = c.Attributes["new_palmclientsupport1.new_ishomeless"].ToString();
                                else
                                    dbIsHomeless = "";

                                // Wrap asterisks around values for better comparisons
                                dbIsHomeless = getMult(dbIsHomeless);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_livingarrangepres"))
                                    dbLivingArrangePres = c.FormattedValues["new_palmclientsupport1.new_livingarrangepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_livingarrangepres"))
                                    dbLivingArrangePres = c.Attributes["new_palmclientsupport1.new_livingarrangepres"].ToString();
                                else
                                    dbLivingArrangePres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_incomepres"))
                                    dbIncomePres = c.FormattedValues["new_palmclientsupport1.new_incomepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_incomepres"))
                                    dbIncomePres = c.Attributes["new_palmclientsupport1.new_incomepres"].ToString();
                                else
                                    dbIncomePres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_interpreterused"))
                                    dbInterpreterUsed = c.FormattedValues["new_palmclientsupport1.new_interpreterused"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_interpreterused"))
                                    dbInterpreterUsed = c.Attributes["new_palmclientsupport1.new_interpreterused"].ToString();
                                else
                                    dbInterpreterUsed = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_sourceref"))
                                    dbSourceRef = c.FormattedValues["new_palmclientsupport1.new_sourceref"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_sourceref"))
                                    dbSourceRef = c.Attributes["new_palmclientsupport1.new_sourceref"].ToString();
                                else
                                    dbSourceRef = "";


                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_primreason"))
                                    dbPrimReason = c.FormattedValues["new_palmclientsupport1.new_primreason"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_primreason"))
                                    dbPrimReason = c.Attributes["new_palmclientsupport1.new_primreason"].ToString();
                                else
                                    dbPrimReason = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_reasons"))
                                    dbReasons = c.FormattedValues["new_palmclientsupport1.new_reasons"]; // multi
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_reasons"))
                                    dbReasons = c.Attributes["new_palmclientsupport1.new_reasons"].ToString();
                                else
                                    dbReasons = "";

                                dbReasons = getMult(dbReasons);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_residentialpres"))
                                    dbResidentialPres = c.FormattedValues["new_palmclientsupport1.new_residentialpres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_residentialpres"))
                                    dbResidentialPres = c.Attributes["new_palmclientsupport1.new_residentialpres"].ToString();
                                else
                                    dbResidentialPres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_dva"))
                                    dbDva = c.FormattedValues["new_palmclientsupport1.new_dva"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_dva"))
                                    dbDva = c.Attributes["new_palmclientsupport1.new_dva"].ToString();
                                else
                                    dbDva = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_careravail"))
                                    dbCarerAvail = c.FormattedValues["new_palmclientsupport1.new_careravail"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_careravail"))
                                    dbCarerAvail = c.Attributes["new_palmclientsupport1.new_careravail"].ToString();
                                else
                                    dbCarerAvail = "";

                                if (c.FormattedValues.Contains("new_palmclientdex2.new_agencyid"))
                                    dbAgencyId = c.FormattedValues["new_palmclientdex2.new_agencyid"];
                                else if (c.Attributes.Contains("new_palmclientdex2.new_agencyid"))
                                    dbAgencyId = c.GetAttributeValue<AliasedValue>("new_palmclientdex2.new_agencyid").Value.ToString();
                                else
                                    dbAgencyId = "";

                                if (c.FormattedValues.Contains("new_palmclientdex2.new_entrydate"))
                                    dbEntryDate = c.FormattedValues["new_palmclientdex2.new_entrydate"];
                                else if (c.Attributes.Contains("new_palmclientdex2.new_entrydate"))
                                    dbEntryDate = c.Attributes["new_palmclientdex2.new_entrydate"].ToString();
                                else
                                    dbEntryDate = "";

                                // Convert date from American format to Australian format
                                dbEntryDate = cleanDateAM(dbEntryDate);

                                if (String.IsNullOrEmpty(dbEntryDate) == false)
                                    varEntryDate = dbEntryDate.Replace("-", "_");
                                else
                                    varEntryDate = "-";

                                // Convert date to required format
                                if (String.IsNullOrEmpty(dbEntryDate) == false)
                                    dbEntryDate = cleanDateE(Convert.ToDateTime(dbEntryDate));

                                // Get required values based on data
                                if (dbInterpreterUsed == "Yes" || dbInterpreterUsed == "yes")
                                    varInterpreterUsed = "true";
                                else
                                    varInterpreterUsed = "false";

                                if (String.IsNullOrEmpty(dbCCPDisType) == false && dbCCPDisType != "-" && dbCCPDisType.ToLower() != "no disability" && dbCCPDisType.ToLower() != "not stated/inadequately described")
                                    varHasDis = "true";
                                else
                                    varHasDis = "false";

                                if (dbIsHomeless.ToLower().IndexOf("short-term or emergency accom") > -1 || dbIsHomeless.ToLower().IndexOf("sleeping rough") > -1)
                                    varIsHomeless = "Yes";
                                else
                                    varIsHomeless = "No";

                                // Ensure postcode is 4 digits
                                if (dbPostcode.Length == 1)
                                    dbPostcode = "000" + dbPostcode;
                                if (dbPostcode.Length == 2)
                                    dbPostcode = "00" + dbPostcode;
                                if (dbPostcode.Length == 3)
                                    dbPostcode = "0" + dbPostcode;

                                varMaxRecip = 0; // Reset

                                //Get Max Recip for this support period
                                foreach (var p in result6.Entities)
                                {
                                    //varTest = "STARTING ATTRIBUTES:\r\n";

                                    //foreach (KeyValuePair<String, Object> attribute in p.Attributes)
                                    //{
                                    //    varTest += (attribute.Key + ": " + attribute.Value + "\r\n");
                                    //}

                                    //varTest += "STARTING FORMATTED:\r\n";

                                    //foreach (KeyValuePair<String, String> value2 in p.FormattedValues)
                                    //{
                                    //    varTest += (value2.Key + ": " + value2.Value + "\r\n");
                                    //}

                                    // We need to get the entity id for the support period field for comparisons
                                    if (p.Attributes.Contains("new_supportperiod_max"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getAlias = p.GetAttributeValue<AliasedValue>("new_supportperiod_max");
                                        getEntity = (EntityReference)getAlias.Value;
                                        dbSupportPeriod_Max = getEntity.Id.ToString();
                                    }
                                    else if (p.FormattedValues.Contains("new_supportperiod_max"))
                                        dbSupportPeriod_Max = p.FormattedValues["new_supportperiod_max"];
                                    else
                                        dbSupportPeriod_Max = "";

                                    // Only do this if support periods are the same
                                    if (dbSupportPeriod_Max == dbPalmClientSupportId)
                                    {
                                        // Get max recip value and convert to integer
                                        if (p.FormattedValues.Contains("new_norecip_max"))
                                            dbNoRecip_Max = p.FormattedValues["new_norecip_max"];
                                        else if (c.Attributes.Contains("new_norecip_max"))
                                            dbNoRecip_Max = p.Attributes["new_norecip_max"].ToString();
                                        else
                                            dbNoRecip_Max = "";

                                        dbNoRecip_Max = cleanString(dbNoRecip_Max, "number");
                                        Int32.TryParse(dbNoRecip_Max, out varCheckInt);

                                        // Set max recip if greater
                                        if (varCheckInt > varMaxRecip)
                                            varMaxRecip = varCheckInt;

                                        break;

                                    } // Same support period
                                } // Max Recip

                                // Only do this if client has not already been processed
                                if (varDoNext == 0)
                                {
                                    //Data Validation
                                    if (String.IsNullOrEmpty(dbLocality) == true)
                                        sbErrorList.AppendLine("(" + dbClient + ") " + dbFirstName + " " + dbSurname + ": Missing Locality<br>");

                                    if (String.IsNullOrEmpty(dbState) == true)
                                        sbErrorList.AppendLine("(" + dbClient + ") " + dbFirstName + " " + dbSurname + ": Missing State<br>");

                                    if (String.IsNullOrEmpty(dbPostcode) == true)
                                        sbErrorList.AppendLine("(" + dbClient + ") " + dbFirstName + " " + dbSurname + ": Missing Postcode<br>");

                                    if (String.IsNullOrEmpty(dbLivingArrangePres) == true)
                                        sbErrorList.AppendLine("(" + dbClient + ") " + dbFirstName + " " + dbSurname + ": Missing Living Arrangements<br>");

                                    if (String.IsNullOrEmpty(dbIncomePres) == true)
                                        sbErrorList.AppendLine("(" + dbClient + ") " + dbFirstName + " " + dbSurname + ": Missing Income<br>");

                                    if (String.IsNullOrEmpty(dbSourceRef) == true)
                                        sbErrorList.AppendLine("(" + dbClient + ") " + dbFirstName + " " + dbSurname + ": Missing Referral Source<br>");

                                    if (String.IsNullOrEmpty(dbPrimReason) == true)
                                        sbErrorList.AppendLine("(" + dbClient + ") " + dbFirstName + " " + dbSurname + ": Missing Primary Presenting Reason<br>");

                                    dbGender = dbGender.ToUpper();
                                    if (dbGender != "MALE" && dbGender != "FEMALE")
                                        dbGender = "NOTSTATED";

                                    // Reset data checkers
                                    varIndigenous = false;
                                    varState = false;
                                    varIncome = false;
                                    varLiving = false;
                                    varSource = false;
                                    varCCP = false;
                                    varResidential = false;
                                    varDV = false;

                                    if (dbCountry == "Don't Know" || dbCountry == "Not Applicable")
                                        dbCountry = "9999";

                                    // Loop through drop down list values
                                    foreach (var d in result3.Entities)
                                    {
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

                                        if (d.FormattedValues.Contains("new_er"))
                                            varDEX = d.FormattedValues["new_er"];
                                        else if (d.Attributes.Contains("new_er"))
                                            varDEX = d.Attributes["new_er"].ToString();
                                        else
                                            varDEX = "";



                                        // Get DEX value based on type and whether the description matches the data
                                        if (varType.ToLower() == "sourceref" && dbSourceRef.ToLower() == varDesc.ToLower())
                                        {

                                            dbSourceRef = varDEX;
                                            varSource = true;
                                        }


                                        if (varType.ToLower() == "indigenous" && dbIndigenous.ToLower() == varDesc.ToLower())
                                        {
                                            dbIndigenous = varDEX;
                                            varIndigenous = true;
                                        }

                                        if (varType.ToLower() == "state" && dbState.ToLower() == varDesc.ToLower())
                                        {
                                            dbState = varDEX;
                                            varState = true;
                                        }

                                        if (varType.ToLower() == "income" && dbIncomePres.ToLower() == varDesc.ToLower())
                                        {
                                            dbIncomePres = varDEX;
                                            varIncome = true;
                                        }

                                        if (varType.ToLower() == "livingarrange" && dbLivingArrangePres.ToLower() == varDesc.ToLower())
                                        {
                                            dbLivingArrangePres = varDEX;
                                            varLiving = true;
                                        }

                                        if (varType.ToLower() == "ccpdis" && dbCCPDisType.ToLower() == varDesc.ToLower())
                                        {
                                            dbCCPDisType = varDEX;
                                            varCCP = true;
                                        }

                                        if (varType.ToLower() == "residential" && dbResidentialPres.ToLower() == varDesc.ToLower())
                                        {
                                            dbResidentialPres = varDEX;
                                            varResidential = true;
                                        }

                                        if (varType.ToLower() == "dva" && dbDva.ToLower() == varDesc.ToLower())
                                        {
                                            dbDva = varDEX;
                                            varDV = true;
                                        }

                                    } // Drop List

                                    // Set defaults for missing data
                                    if (varSource == false)
                                        dbSourceRef = "NOTSTATED";

                                    if (varIndigenous == false)
                                        dbIndigenous = "NOTSTATED";

                                    if (varState == false)
                                        dbState = "VIC";

                                    if (varIncome == false)
                                        dbIncomePres = "NOTSTATED";

                                    if (varLiving == false)
                                        dbLivingArrangePres = "NOTSTATED";

                                    if (varCCP == false)
                                        dbCCPDisType = "NOTSTATED";

                                    if (varResidential == false)
                                        dbResidentialPres = "NOTSTATED";

                                    if (varDV == false)
                                        dbDva = "NODVA";

                                    // Convert to valid values
                                    if (dbConsentFuture == "Yes" || dbConsentFuture == "true")
                                        dbConsentFuture = "true";
                                    else
                                        dbConsentFuture = "false";

                                    if (dbConsentRelease == "Yes" || dbConsentRelease == "true")
                                        dbConsentRelease = "true";
                                    else
                                        dbConsentRelease = "false";


                                    // Create client part of extract
                                    sbClientList.AppendLine("   <Client>");

                                    //Mandatory String
                                    sbClientList.AppendLine("      <ClientId>" + dbClient + "</ClientId>");

                                    //Optional String - Mandatory if consent false
                                    if (dbDexSlk != "-")
                                        sbClientList.AppendLine("      <Slk>" + dbDexSlk + "</Slk>");

                                    if (dbDexSlk == "-" && dbConsentRelease == "false")
                                        sbErrorList.AppendLine("(" + dbClient + ") " + dbFirstName + " " + dbSurname + " has not given consent and does not have an SLK<br>");

                                    //Mandatory True / False
                                    if (String.IsNullOrEmpty(dbConsentRelease) == false)
                                        sbClientList.AppendLine("      <ConsentToProvideDetails>" + dbConsentRelease + "</ConsentToProvideDetails>");
                                    else
                                        sbClientList.AppendLine("      <ConsentToProvideDetails>false</ConsentToProvideDetails>");

                                    if (String.IsNullOrEmpty(dbConsentFuture) == false)
                                        sbClientList.AppendLine("      <ConsentedForFutureContacts>" + dbConsentFuture + "</ConsentedForFutureContacts>");
                                    else
                                        sbClientList.AppendLine("      <ConsentedForFutureContacts>false</ConsentedForFutureContacts>");

                                    //Conditional String - Not allowed if Consent False
                                    if (dbConsentRelease == "true")
                                    {
                                        sbClientList.AppendLine("      <GivenName>" + dbFirstName + "</GivenName>");
                                        sbClientList.AppendLine("      <FamilyName>" + dbSurname + "</FamilyName>");
                                    }

                                    //Always False
                                    sbClientList.AppendLine("      <IsUsingPsuedonym>false</IsUsingPsuedonym>");

                                    //Mandatory YYYY-MM-DD
                                    sbClientList.AppendLine("      <BirthDate>" + dbDob + "</BirthDate>");
                                    //Mandatory True / False (if true DOB must be YYYY-01-01)
                                    sbClientList.AppendLine("      <IsBirthDateAnEstimate>" + dbDobEst + "</IsBirthDateAnEstimate>");

                                    //Mandatory Reference Data
                                    sbClientList.AppendLine("      <GenderCode>" + dbGender + "</GenderCode>");
                                    sbClientList.AppendLine("      <CountryOfBirthCode>" + dbCountry + "</CountryOfBirthCode>");
                                    sbClientList.AppendLine("      <LanguageSpokenAtHomeCode>" + dbLanguage + "</LanguageSpokenAtHomeCode>");
                                    sbClientList.AppendLine("      <AboriginalOrTorresStraitIslanderOriginCode>" + dbIndigenous + "</AboriginalOrTorresStraitIslanderOriginCode>");

                                    sbClientList.AppendLine("      <HasDisabilities>" + varHasDis + "</HasDisabilities>"); //Mandatory True / False

                                    //XML if has disabilities true
                                    if (varHasDis == "true")
                                    {
                                        sbClientList.AppendLine("      <Disabilities>");
                                        sbClientList.AppendLine("         <DisabilityCode>" + dbCCPDisType + "</DisabilityCode>"); //Mandatory Reference
                                        sbClientList.AppendLine("      </Disabilities>");
                                    }


                                    //if (varDoCHSP == true)
                                    //{
                                    sbClientList.AppendLine("      <AccommodationTypeCode>" + dbResidentialPres + "</AccommodationTypeCode>");
                                    sbClientList.AppendLine("      <DVACardStatusCode>" + dbDva + "</DVACardStatusCode>");


                                    if (dbCarerAvail == "Yes" || dbCarerAvail == "Has A Carer")
                                        sbClientList.AppendLine("      <HasCarer>true</HasCarer>");
                                    else
                                        sbClientList.AppendLine("      <HasCarer>false</HasCarer>");
                                    //}


                                    //XML for address
                                    sbClientList.AppendLine("      <ResidentialAddress>");

                                    //Optional String - not using
                                    //sbClientList.AppendLine("         <AddressLine1>[...]</AddressLine1>");
                                    //sbClientList.AppendLine("         <AddressLine2>[...]</AddressLine2>");

                                    //Mandatory String, reference and 4 character string
                                    sbClientList.AppendLine("         <Suburb>" + dbLocality + "</Suburb>");
                                    sbClientList.AppendLine("         <StateCode>" + dbState + "</StateCode>");
                                    sbClientList.AppendLine("         <Postcode>" + dbPostcode + "</Postcode>");
                                    sbClientList.AppendLine("      </ResidentialAddress>");

                                    //Optional True / False
                                    sbClientList.AppendLine("      <HomelessIndicatorCode>" + varIsHomeless + "</HomelessIndicatorCode>");

                                    //Optional Reference
                                    sbClientList.AppendLine("      <HouseholdCompositionCode>" + dbLivingArrangePres + "</HouseholdCompositionCode>");
                                    //Mandatory Reference
                                    sbClientList.AppendLine("      <MainSourceOfIncomeCode>" + dbIncomePres + "</MainSourceOfIncomeCode>");

                                    //sbClientList.AppendLine("      <IncomeFrequencyCode>[...]</IncomeFrequencyCode>"); //Optional Reference - not using
                                    //sbClientList.AppendLine("      <IncomeAmount>[...]</IncomeAmount>"); //Optional Number - not using

                                    // Add year of arrival
                                    if (dbYearArrival != "-" && dbYearArrival != "0" && String.IsNullOrEmpty(dbYearArrival) == false)
                                    {
                                        sbClientList.AppendLine("      <FirstArrivalYear>" + dbYearArrival + "</FirstArrivalYear>"); //Optional Number - greater/equal DOB, less than Today
                                        sbClientList.AppendLine("      <FirstArrivalMonth>January</FirstArrivalMonth>"); //Mandatory - December
                                    }

                                    //sbClientList.AppendLine("      <MigrationVisaCategoryCode>[...]</MigrationVisaCategoryCode>"); //Optional Reference - not using
                                    //sbClientList.AppendLine("      <AncestryCode>[...]</AncestryCode>"); //Optional Reference - not using

                                    sbClientList.AppendLine("   </Client>");

                                } // Client already processed


                                //Reset duplicate support period test
                                varDoNext3 = 0;
                                if (varCaseNumber.IndexOf("*" + dbPalmClientSupportId + "*") > -1)
                                    varDoNext3 = 1;

                                varCaseNumber += "*" + dbPalmClientSupportId + "*";

                                // Only do if support period not already processed
                                if (varDoNext3 == 0)
                                {

                                    // Build Case Here
                                    sbCaseList.AppendLine("   <Case>");

                                    //Mandatory String
                                    sbCaseList.AppendLine("      <CaseId>" + varPalmClientSupportId + "_" + varExt + "</CaseId>");

                                    //ToDo: Mandatory Integer
                                    sbCaseList.AppendLine("      <OutletActivityId>" + varOutletActivityId + "</OutletActivityId>");

                                    //Mandatory Integer
                                    if (varMaxRecip > 0)
                                        sbCaseList.AppendLine("      <TotalNumberOfUnidentifiedClients>" + varMaxRecip + "</TotalNumberOfUnidentifiedClients>");
                                    else
                                        sbCaseList.AppendLine("      <TotalNumberOfUnidentifiedClients>0</TotalNumberOfUnidentifiedClients>");

                                    //Optional
                                    sbCaseList.AppendLine("      <CaseClients>");

                                    sbCaseList.AppendLine("         <CaseClient>");

                                    sbCaseList.AppendLine("            <ClientId>" + dbClient + "</ClientId>");


                                    //This is an extremely hacky solution that is to get around an issue on the 6/05/2022 - one day I might be able to fix this. Today is not that day.
                                    if(dbSourceRef.ToUpper() == "OTHER COMMUNITY-BASED SERVICE")
                                    {
                                        dbSourceRef = "COMMUNITY";
                                    }

                                    sbCaseList.AppendLine("            <ReferralSourceCode>" + dbSourceRef.ToUpper() + "</ReferralSourceCode>");     

                                    // Append reasons for assistance if reasons not empty
                                    if (String.IsNullOrEmpty(dbReasons) == false)
                                    {

                                        sbCaseList.AppendLine("            <ReasonsForAssistance>");
                                        varReasonStr = "";
                                        // Loop through drop down list and get integer for primary reason
                                        foreach (var d in result3.Entities)
                                        {

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

                                            if (d.FormattedValues.Contains("new_er"))
                                                varDEX = d.FormattedValues["new_er"];
                                            else if (d.Attributes.Contains("new_er"))
                                                varDEX = d.Attributes["new_er"].ToString();
                                            else
                                                varDEX = "";

                                            // Add reason
                                            if (varDesc.ToLower() == dbPrimReason.ToLower() && varType == "reasons")
                                            {
                                                sbCaseList.AppendLine("               <ReasonForAssistance>");
                                                sbCaseList.AppendLine("                  <ReasonForAssistanceCode>" + varDEX + "</ReasonForAssistanceCode>");
                                                sbCaseList.AppendLine("                  <IsPrimary>true</IsPrimary>");
                                                sbCaseList.AppendLine("               </ReasonForAssistance>");

                                                varReasonStr += "*" + varDEX + "*";

                                                break;
                                            }

                                        } //k Loop


                                        // Loop through drop down list and get reasons for presenting
                                        foreach (var d in result3.Entities)
                                        {

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

                                            if (d.FormattedValues.Contains("new_er"))
                                                varDEX = d.FormattedValues["new_er"];
                                            else if (d.Attributes.Contains("new_er"))
                                                varDEX = d.Attributes["new_er"].ToString();
                                            else
                                                varDEX = "";

                                            // Add reason
                                            varTest = dbReasons;
                                            //sbCaseList.AppendLine("<DebugSection1>Greater than -1? " + dbReasons.IndexOf("*" + varDesc + "*") + " .. Not equal to " + dbPrimReason.ToLower() + "? " + varDesc.ToLower() + " .. Equal to -1?" + varReasonStr.IndexOf("*" + varDEX + "*") + " .. reasons?" + varType);
                                            if (dbReasons.IndexOf("*" + varDesc + "*") > -1 && varDesc.ToLower() != dbPrimReason.ToLower() && varReasonStr.IndexOf("*" + varDEX + "*") == -1 && varType == "reasons")
                                            {
                                                sbCaseList.AppendLine("               <ReasonForAssistance>");
                                                sbCaseList.AppendLine("                  <ReasonForAssistanceCode>" + varDEX + "</ReasonForAssistanceCode>");
                                                sbCaseList.AppendLine("                  <IsPrimary>false</IsPrimary>");
                                                sbCaseList.AppendLine("               </ReasonForAssistance>");

                                                varReasonStr += "*" + varDEX + "*";
                                            }

                                        } //j Loop

                                        sbCaseList.AppendLine("            </ReasonsForAssistance>");

                                    } //Reasons

                                    sbCaseList.AppendLine("         </CaseClient>");

                                    sbCaseList.AppendLine("      </CaseClients>");

                                    sbCaseList.AppendLine("   </Case>");

                                    //if (dbClient == "PLM31076")
                                    //    varTest += "Support Id: " + dbPalmClientSupportId + "\r\n";

                                    // Loop through data from DEX table
                                    foreach (var x in result2.Entities)
                                    {
                                        //varTest = "STARTING ATTRIBUTES:\r\n";

                                        //foreach (KeyValuePair<String, Object> attribute in s.Attributes)
                                        //{
                                        //    varTest += (attribute.Key + ": " + attribute.Value + "\r\n");
                                        //}

                                        //varTest += "STARTING FORMATTED:\r\n";

                                        //foreach (KeyValuePair<String, String> value in s.FormattedValues)
                                        //{
                                        //    varTest += (value.Key + ": " + value.Value + "\r\n");
                                        //}

                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (x.FormattedValues.Contains("new_palmclientdexid"))
                                            dbPalmClientDexId = x.FormattedValues["new_palmclientdexid"];
                                        else if (x.Attributes.Contains("new_palmclientdexid"))
                                            dbPalmClientDexId = x.Attributes["new_palmclientdexid"].ToString();
                                        else
                                            dbPalmClientDexId = "";

                                        // Remove dashes
                                        varPalmClientDexId = "";
                                        if (String.IsNullOrEmpty(dbPalmClientDexId) == false)
                                            varPalmClientDexId = dbPalmClientDexId.Replace("-", "");

                                        if (x.FormattedValues.Contains("new_entrydate"))
                                            dbEntryDate2 = x.FormattedValues["new_entrydate"];
                                        else if (x.Attributes.Contains("new_entrydate"))
                                            dbEntryDate2 = x.Attributes["new_entrydate"].ToString();
                                        else
                                            dbEntryDate2 = "";

                                        // Convert date from American format to Australian format
                                        dbEntryDate2 = cleanDateAM(dbEntryDate2);

                                        if (String.IsNullOrEmpty(dbEntryDate2) == false)
                                            varEntryDate2 = dbEntryDate2.Replace("-", "_");
                                        else
                                            varEntryDate2 = "-";

                                        if (x.FormattedValues.Contains("new_consentfuture"))
                                            dbConsentFuture = x.FormattedValues["new_consentfuture"];
                                        else if (x.Attributes.Contains("new_consentfuture"))
                                            dbConsentFuture = x.Attributes["new_consentfuture"].ToString();
                                        else
                                            dbConsentFuture = "";

                                        if (x.FormattedValues.Contains("new_consentrelease"))
                                            dbConsentRelease = x.FormattedValues["new_consentrelease"];
                                        else if (x.Attributes.Contains("new_consentrelease"))
                                            dbConsentRelease = x.Attributes["new_consentfuture"].ToString();
                                        else
                                            dbConsentRelease = "";

                                        if (x.FormattedValues.Contains("new_norecip"))
                                            dbNoRecip = x.FormattedValues["new_norecip"];
                                        else if (x.Attributes.Contains("new_norecip"))
                                            dbNoRecip = x.Attributes["new_norecip"].ToString();
                                        else
                                            dbNoRecip = "";

                                        dbNoRecip = cleanString(dbNoRecip, "number");
                                        Int32.TryParse(dbNoRecip, out varCheckInt);

                                        varNoRecip = varCheckInt;
                                        if (varNoRecip > varGetRecip)
                                            varGetRecip = varNoRecip;

                                        if (x.FormattedValues.Contains("new_hours"))
                                            dbHours = x.FormattedValues["new_hours"];
                                        else if (x.Attributes.Contains("new_hours"))
                                            dbHours = x.Attributes["new_hours"].ToString();
                                        else
                                            dbHours = "";

                                        dbHours = cleanString(dbHours, "double");

                                        Double.TryParse(dbHours, out varCheckDouble);
                                        dbHours = (int)(varCheckDouble * 60) + "";

                                        if (x.FormattedValues.Contains("new_servtype"))
                                            dbServType = x.FormattedValues["new_servtype"];
                                        else if (x.Attributes.Contains("new_servtype"))
                                            dbServType = x.Attributes["new_servtype"].ToString();
                                        else
                                            dbServType = "";

                                        // We need to get the entity id for the support period field for comparisons
                                        if (x.Attributes.Contains("new_supportperiod"))
                                        {
                                            // Get the entity id for the client using the entity reference object
                                            getEntity = (EntityReference)x.Attributes["new_supportperiod"];
                                            dbSupportPeriod = getEntity.Id.ToString();
                                        }
                                        else if (x.FormattedValues.Contains("new_supportperiod"))
                                            dbSupportPeriod = x.FormattedValues["new_supportperiod"];
                                        else
                                            dbSupportPeriod = "";

                                        // We need to get the entity id for the agency id field for comparisons
                                        if (x.Attributes.Contains("new_agencyid"))
                                        {
                                            // Get the entity id for the client using the entity reference object
                                            getEntity = (EntityReference)x.Attributes["new_agencyid"];
                                            dbAgencyId2 = getEntity.Id.ToString();
                                        }
                                        else if (x.FormattedValues.Contains("new_agencyid"))
                                            dbAgencyId2 = x.FormattedValues["new_agencyid"];
                                        else
                                            dbAgencyId2 = "";

                                        // Only do if the support periods match
                                        if (dbSupportPeriod == dbPalmClientSupportId)
                                        {
                                            // Check to see if session has already been processed
                                            varDoNext2 = 0;
                                            if (varSessionNumber.IndexOf("*" + dbPalmClientDexId + "*") > -1)
                                                varDoNext2 = 1;

                                            varSessionNumber += "*" + dbPalmClientDexId + "*";

                                            //if (dbClient == "PLM31076")
                                            //    varTest += "DEX Id: " + dbPalmClientDexId + "\r\n";

                                            // Only do if session not processed
                                            if (varDoNext2 == 0)
                                            {
                                                // Build Activity here

                                                // Session assessments for CHSP                          
                                                if (varDoCHSP == true)
                                                {
                                                    // Set string to any value to prevent index search errors
                                                    varPreGroup = "a";
                                                    varPostGroup = "a";

                                                    // Reset session variables
                                                    sbSessionAssessment2List.Length = 0;
                                                    varSeeClient3 = false;

                                                    //ER Score loop
                                                    foreach (var s in result8.Entities)
                                                    {
                                                        // Process the data as follows:
                                                        // If there is a formatted value for the field, use it
                                                        // Otherwise if there is a literal value for the field, use it
                                                        // Otherwise the value wasn't returned so set as nothing
                                                        if (s.FormattedValues.Contains("new_palmclientdexid"))
                                                            dbPalmClientDexId2 = s.FormattedValues["new_palmclientdexid"];
                                                        else if (s.Attributes.Contains("new_palmclientdexid"))
                                                            dbPalmClientDexId2 = s.Attributes["new_palmclientdexid"].ToString();
                                                        else
                                                            dbPalmClientDexId2 = "";

                                                        if (s.FormattedValues.Contains("new_palmclientdexscores1.new_palmclientdexscoresid"))
                                                            dbPalmClientDexScoresId = s.FormattedValues["new_palmclientdexscores1.new_palmclientdexscoresid"];
                                                        else if (s.Attributes.Contains("new_palmclientdexscores1.new_palmclientdexscoresid"))
                                                            dbPalmClientDexScoresId = s.GetAttributeValue<AliasedValue>("new_palmclientdexscores1.new_palmclientdexscoresid").Value.ToString();
                                                        else
                                                            dbPalmClientDexScoresId = "";

                                                        if (s.FormattedValues.Contains("new_palmclientdexscores1.new_score"))
                                                            dbScore = s.FormattedValues["new_palmclientdexscores1.new_score"];
                                                        else if (s.Attributes.Contains("new_palmclientdexscores1.new_score"))
                                                            dbScore = s.Attributes["new_palmclientdexscores1.new_score"].ToString();
                                                        else
                                                            dbScore = "";

                                                        if (s.FormattedValues.Contains("new_palmclientdexscores1.new_section"))
                                                            dbSection = s.FormattedValues["new_palmclientdexscores1.new_section"];
                                                        else if (s.Attributes.Contains("new_palmclientdexscores1.new_section"))
                                                            dbSection = s.Attributes["new_palmclientdexscores1.new_section"].ToString();
                                                        else
                                                            dbSection = "";

                                                        // Get the section part of the field
                                                        if (String.IsNullOrEmpty(dbSection) == false)
                                                        {
                                                            dbSection = dbSection.ToLower();
                                                            if (dbSection.IndexOf(" - ") > -1)
                                                                varSection = dbSection.Substring(dbSection.IndexOf(" - ") + 3, dbSection.Length - dbSection.IndexOf(" - ") - 3);
                                                            else
                                                                varSection = "...";
                                                        }
                                                        else
                                                        {
                                                            dbSection = "...";
                                                            varSection = "...";
                                                        }

                                                        if (s.FormattedValues.Contains("new_palmclientdexscores1.new_dexrecord"))
                                                            dbDexRecord = s.FormattedValues["new_palmclientdexscores1.new_dexrecord"];
                                                        else if (s.Attributes.Contains("new_palmclientdexscores1.new_dexrecord"))
                                                            dbDexRecord = s.GetAttributeValue<AliasedValue>("new_palmclientdexscores1.new_dexrecord").Value.ToString();
                                                        else
                                                            dbDexRecord = "";

                                                        if (s.FormattedValues.Contains("new_palmclientdexscores1.new_thetype"))
                                                            dbTheType = s.FormattedValues["new_palmclientdexscores1.new_thetype"];
                                                        else if (s.Attributes.Contains("new_palmclientdexscores1.new_thetype"))
                                                            dbTheType = s.Attributes["new_palmclientdexscores1.new_thetype"].ToString();
                                                        else
                                                            dbTheType = "";

                                                        // Only do if the dex record ids match
                                                        if (dbPalmClientDexId == dbPalmClientDexId2)
                                                        {
                                                            // Score group section
                                                            if (dbSection.IndexOf("scoregroup") > -1)
                                                            {
                                                                //Get the values from the drop down tables
                                                                foreach (var d in result3.Entities)
                                                                {
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

                                                                    if (d.FormattedValues.Contains("new_er"))
                                                                        varDEX = d.FormattedValues["new_er"];
                                                                    else if (d.Attributes.Contains("new_er"))
                                                                        varDEX = d.Attributes["new_er"].ToString();
                                                                    else
                                                                        varDEX = "";

                                                                    // Append the section and score to the pre / post group section accordingly
                                                                    if (varSection.ToLower() == varDesc.ToLower())
                                                                    {
                                                                        dbSection = dbSection.Replace(varDesc.ToLower(), varDEX);
                                                                        dbSection = dbSection.Replace("scoregroup - ", "");

                                                                        if (dbTheType.ToLower() == "pre")
                                                                            varPreGroup += "," + dbSection + dbScore;
                                                                        if (dbTheType.ToLower() == "post")
                                                                            varPostGroup += "," + dbSection + dbScore;
                                                                        break;
                                                                    }

                                                                } //drop down Loop
                                                            }

                                                        } //Same client

                                                        varSeeClient3 = true; // record found

                                                    } //Scores Loop

                                                    // If the pre group was appended to above, add to extract
                                                    if (varPreGroup != "a")
                                                    {

                                                        sbSessionAssessment2List.AppendLine("         <Assessment>");

                                                        sbSessionAssessment2List.AppendLine("            <ScoreTypeCode>GROUP</ScoreTypeCode>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("            <AssessmentPhaseCode>PRE</AssessmentPhaseCode>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                        // Loop through string to get scores
                                                        value = varPreGroup.Split(',');

                                                        foreach (string dc in value)
                                                        {
                                                            if (dc != "a")
                                                                sbSessionAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                        }

                                                        sbSessionAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("         </Assessment>");

                                                    } //"a"

                                                    // If the post group was appended to above, add to extract
                                                    if (varPostGroup != "a")
                                                    {

                                                        sbSessionAssessment2List.AppendLine("         <Assessment>");

                                                        sbSessionAssessment2List.AppendLine("            <ScoreTypeCode>GROUP</ScoreTypeCode>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("            <AssessmentPhaseCode>POST</AssessmentPhaseCode>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                        // Loop through string to get scores
                                                        value = varPostGroup.Split(',');

                                                        foreach (string dc in value)
                                                        {
                                                            if (dc != "a")
                                                                sbSessionAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                        }

                                                        sbSessionAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("         </Assessment>");

                                                    } //"a"

                                                    // If the assessment 2 list was appended to above, add the session assessment to the extract
                                                    if (sbSessionAssessment2List.Length > 0)
                                                    {

                                                        sbSessionAssessmentList.AppendLine("   <SessionAssessment>");
                                                        sbSessionAssessmentList.AppendLine("      <CaseId>" + varPalmClientSupportId + "_" + varExt + "</CaseId>"); //Mandatory
                                                        sbSessionAssessmentList.AppendLine("      <SessionId>" + varPalmClientDexId + "_" + varExt + "</SessionId>"); //Mandatory

                                                        sbSessionAssessmentList.AppendLine("      <Assessments>");

                                                        sbSessionAssessmentList.AppendLine(sbSessionAssessment2List.ToString());

                                                        sbSessionAssessmentList.AppendLine("      </Assessments>");
                                                        sbSessionAssessmentList.AppendLine("   </SessionAssessment>");

                                                    }

                                                } //CHSP


                                                // Client Assessments
                                                varPreCirc = "a"; // Create default data
                                                varPostCirc = "a"; // Create default data
                                                varPreGoal = "a"; // Create default data
                                                varPostGoal = "a"; // Create default data
                                                varSatisfaction = "a"; // Create default data

                                                // Reset assessment variables
                                                sbClientAssessment2List.Length = 0;
                                                varSeeClient3 = false;


                                                //ER Score loop
                                                foreach (var s in result8.Entities)
                                                {
                                                    // Process the data as follows:
                                                    // If there is a formatted value for the field, use it
                                                    // Otherwise if there is a literal value for the field, use it
                                                    // Otherwise the value wasn't returned so set as nothing
                                                    if (s.FormattedValues.Contains("new_palmclientdexid"))
                                                        dbPalmClientDexId2 = s.FormattedValues["new_palmclientdexid"];
                                                    else if (s.Attributes.Contains("new_palmclientdexid"))
                                                        dbPalmClientDexId2 = s.Attributes["new_palmclientdexid"].ToString();
                                                    else
                                                        dbPalmClientDexId2 = "";

                                                    if (s.FormattedValues.Contains("new_palmclientdexscores1.new_palmclientdexscoresid"))
                                                        dbPalmClientDexScoresId = s.FormattedValues["new_palmclientdexscores1.new_palmclientdexscoresid"];
                                                    else if (s.Attributes.Contains("new_palmclientdexscores1.new_palmclientdexscoresid"))
                                                        dbPalmClientDexScoresId = s.GetAttributeValue<AliasedValue>("new_palmclientdexscores1.new_palmclientdexscoresid").Value.ToString();
                                                    else
                                                        dbPalmClientDexScoresId = "";

                                                    if (s.FormattedValues.Contains("new_palmclientdexscores1.new_score"))
                                                        dbScore = s.FormattedValues["new_palmclientdexscores1.new_score"];
                                                    else if (s.Attributes.Contains("new_palmclientdexscores1.new_score"))
                                                        dbScore = s.Attributes["new_palmclientdexscores1.new_score"].ToString();
                                                    else
                                                        dbScore = "";

                                                    if (s.FormattedValues.Contains("new_palmclientdexscores1.new_section"))
                                                        dbSection = s.FormattedValues["new_palmclientdexscores1.new_section"];
                                                    else if (s.Attributes.Contains("new_palmclientdexscores1.new_section"))
                                                        dbSection = s.Attributes["new_palmclientdexscores1.new_section"].ToString();
                                                    else
                                                        dbSection = "";

                                                    // Get the section part of the field
                                                    if (String.IsNullOrEmpty(dbSection) == false)
                                                    {
                                                        dbSection = dbSection.ToLower();
                                                        if (dbSection.IndexOf(" - ") > -1)
                                                            varSection = dbSection.Substring(dbSection.IndexOf(" - ") + 3, dbSection.Length - dbSection.IndexOf(" - ") - 3);
                                                        else
                                                            varSection = "...";
                                                    }
                                                    else
                                                    {
                                                        dbSection = "...";
                                                        varSection = "...";
                                                    }

                                                    if (s.FormattedValues.Contains("new_palmclientdexscores1.new_dexrecord"))
                                                        dbDexRecord = s.FormattedValues["new_palmclientdexscores1.new_dexrecord"];
                                                    else if (s.Attributes.Contains("new_palmclientdexscores1.new_dexrecord"))
                                                        dbDexRecord = s.GetAttributeValue<AliasedValue>("new_palmclientdexscores1.new_dexrecord").Value.ToString();
                                                    else
                                                        dbDexRecord = "";

                                                    if (s.FormattedValues.Contains("new_palmclientdexscores1.new_thetype"))
                                                        dbTheType = s.FormattedValues["new_palmclientdexscores1.new_thetype"];
                                                    else if (s.Attributes.Contains("new_palmclientdexscores1.new_thetype"))
                                                        dbTheType = s.Attributes["new_palmclientdexscores1.new_thetype"].ToString();
                                                    else
                                                        dbTheType = "";

                                                    // We need to get the entity id for the support period field for comparisons
                                                    if (s.Attributes.Contains("new_supportperiod"))
                                                    {
                                                        // Get the entity id for the client using the entity reference object
                                                        getEntity = (EntityReference)s.Attributes["new_supportperiod"];
                                                        dbSupportPeriod2 = getEntity.Id.ToString();
                                                    }
                                                    else if (s.FormattedValues.Contains("new_supportperiod"))
                                                        dbSupportPeriod2 = s.FormattedValues["new_supportperiod"];
                                                    else
                                                        dbSupportPeriod2 = "";

                                                    //if (dbClient == "PLM31076")
                                                    //    varTest += dbSupportPeriod2 + " " + dbPalmClientDexId2 + " " + dbDexRecord + " " + dbTheType + " " + dbSection + " " + dbScore + "\r\n";

                                                    // Only process if the DEX ids match
                                                    if (dbPalmClientDexId == dbPalmClientDexId2)
                                                    {
                                                        if (dbSection.IndexOf("scorecircumstances") > -1 || dbSection.IndexOf("scoregoals") > -1 || dbSection.IndexOf("scoresatisfaction") > -1)
                                                        {
                                                            //Get the values from the drop down tables
                                                            foreach (var d in result3.Entities)
                                                            {
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

                                                                if (d.FormattedValues.Contains("new_er"))
                                                                    varDEX = d.FormattedValues["new_er"];
                                                                else if (d.Attributes.Contains("new_er"))
                                                                    varDEX = d.Attributes["new_er"].ToString();
                                                                else
                                                                    varDEX = "";

                                                                // Process if the section matches and the type is valid
                                                                if (varSection.ToLower() == varDesc.ToLower() && (varType == "scorecircumstances" || varType == "scoregoals" || varType == "scoresatisfaction"))
                                                                {
                                                                    //if (dbClient == "PLM31076")
                                                                    //    varTest += "*" + varSection + "* *" + varType + " " + varDesc + "* *" + varDEX + "*\r\n";

                                                                    dbSection = dbSection.Replace(varDesc.ToLower(), varDEX);

                                                                    // Append the section and score to the pre / post goals section accordingly
                                                                    if (dbSection.IndexOf("scoregoals") > -1)
                                                                    {
                                                                        dbSection = dbSection.Replace("scoregoals - ", "");
                                                                        if (dbTheType.ToLower() == "pre")
                                                                        {
                                                                            varPreGoal += "," + dbSection + dbScore;
                                                                        }
                                                                        if (dbTheType.ToLower() == "post")
                                                                        {
                                                                            varPostGoal += "," + dbSection + dbScore;
                                                                        }
                                                                    }
                                                                    // Append the section and score to the pre / post circumstances section accordingly
                                                                    if (dbSection.IndexOf("scorecircumstances") > -1)
                                                                    {
                                                                        dbSection = dbSection.Replace("scorecircumstances - ", "");
                                                                        if (dbTheType.ToLower() == "pre")
                                                                        {
                                                                            varPreCirc += "," + dbSection + dbScore;
                                                                        }
                                                                        if (dbTheType.ToLower() == "post")
                                                                        {
                                                                            varPostCirc += "," + dbSection + dbScore;
                                                                        }
                                                                    }
                                                                    // Append the section and score to the post satisfaction section accordingly
                                                                    if (dbSection.IndexOf("scoresatisfaction") > -1)
                                                                    {
                                                                        dbSection = dbSection.Replace("scoresatisfaction - ", "");
                                                                        if (dbTheType.ToLower() == "post")
                                                                        {
                                                                            varSatisfaction += "," + dbSection + dbScore;
                                                                        }
                                                                    }
                                                                    break;
                                                                }

                                                            } //drop down Loop
                                                        }

                                                    } //Same client

                                                    varSeeClient3 = true; // Client record found

                                                } //Scores Loop


                                                // If the pre circ was appended to above, add to extract
                                                if (varPreCirc != "a")
                                                {

                                                    sbClientAssessment2List.AppendLine("         <Assessment>");

                                                    sbClientAssessment2List.AppendLine("            <ScoreTypeCode>CIRCUMSTANCES</ScoreTypeCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <AssessmentPhaseCode>PRE</AssessmentPhaseCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                    // Loop through string to get scores
                                                    value = varPreCirc.Split(',');

                                                    foreach (string dc in value)
                                                    {
                                                        if (dc != "a")
                                                            sbClientAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                    }

                                                    sbClientAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("         </Assessment>");

                                                } //"a"

                                                // If the post circ was appended to above, add to extract
                                                if (varPostCirc != "a")
                                                {

                                                    sbClientAssessment2List.AppendLine("         <Assessment>");

                                                    sbClientAssessment2List.AppendLine("            <ScoreTypeCode>CIRCUMSTANCES</ScoreTypeCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <AssessmentPhaseCode>POST</AssessmentPhaseCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                    // Loop through string to get scores
                                                    value = varPostCirc.Split(',');

                                                    foreach (string dc in value)
                                                    {
                                                        if (dc != "a")
                                                            sbClientAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                    }

                                                    sbClientAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("         </Assessment>");

                                                } //"a"

                                                // If the pre goal was appended to above, add to extract
                                                if (varPreGoal != "a")
                                                {

                                                    sbClientAssessment2List.AppendLine("         <Assessment>");

                                                    sbClientAssessment2List.AppendLine("            <ScoreTypeCode>GOALS</ScoreTypeCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <AssessmentPhaseCode>PRE</AssessmentPhaseCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                    // Loop through string to get scores
                                                    value = varPreGoal.Split(',');

                                                    foreach (string dc in value)
                                                    {
                                                        if (dc != "a")
                                                            sbClientAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                    }

                                                    sbClientAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("         </Assessment>");

                                                } //"a"

                                                // If the post goal was appended to above, add to extract
                                                if (varPostGoal != "a")
                                                {

                                                    sbClientAssessment2List.AppendLine("         <Assessment>");

                                                    sbClientAssessment2List.AppendLine("            <ScoreTypeCode>GOALS</ScoreTypeCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <AssessmentPhaseCode>POST</AssessmentPhaseCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                    // Loop through string to get scores
                                                    value = varPostGoal.Split(',');

                                                    foreach (string dc in value)
                                                    {
                                                        if (dc != "a")
                                                            sbClientAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                    }

                                                    sbClientAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("         </Assessment>");

                                                } //"a"

                                                // If the satisfaction was appended to above, add to extract
                                                if (varSatisfaction != "a")
                                                {

                                                    sbClientAssessment2List.AppendLine("         <Assessment>");

                                                    sbClientAssessment2List.AppendLine("            <ScoreTypeCode>SATISFACTION</ScoreTypeCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <AssessmentPhaseCode>POST</AssessmentPhaseCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                    // Loop through string to get scores
                                                    value = varSatisfaction.Split(',');

                                                    foreach (string dc in value)
                                                    {
                                                        if (dc != "a")
                                                            sbClientAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                    }

                                                    sbClientAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("         </Assessment>");

                                                } //"a"

                                                // If the assessment 2 list was appended to above, add the client assessment to the extract
                                                if (sbClientAssessment2List.Length > 0)
                                                {

                                                    sbClientAssessmentList.AppendLine("   <ClientAssessment>");
                                                    sbClientAssessmentList.AppendLine("      <ClientId>" + dbClient + "</ClientId>"); //Mandatory
                                                    sbClientAssessmentList.AppendLine("      <CaseId>" + varPalmClientSupportId + "_" + varExt + "</CaseId>"); //Mandatory
                                                    sbClientAssessmentList.AppendLine("      <SessionId>" + varPalmClientDexId + "_" + varExt + "</SessionId>"); //Mandatory

                                                    sbClientAssessmentList.AppendLine("      <Assessments>");

                                                    sbClientAssessmentList.AppendLine(sbClientAssessment2List.ToString());

                                                    sbClientAssessmentList.AppendLine("      </Assessments>");
                                                    sbClientAssessmentList.AppendLine("   </ClientAssessment>");

                                                }

                                                // Get the service type id
                                                varServTypeId = "1";
                                                if (varDoCHSP == true)
                                                    varServTypeId = "192";

                                                //Get the assistance type and convert to servicetype id - result3 = dropdownlist.
                                                foreach (var d in result3.Entities)
                                                {
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

                                                    if (d.FormattedValues.Contains("new_er"))
                                                        varDEX = d.FormattedValues["new_er"];
                                                    else if (d.Attributes.Contains("new_er"))
                                                        varDEX = d.Attributes["new_er"].ToString();
                                                    else
                                                        varDEX = "";

                                                    if (varDesc.ToLower() == dbServType.ToLower() && varType == "erassist")
                                                    {
                                                        varServTypeId = varDEX; //20 is - accom, bonds, ooh bond debt, rent arrears, rent in advance.
                                                                                //varDebug = varDesc;
                                                        break;
                                                    }

                                                } //k Loop

                                                // Build Session here                              
                                                sbSessionList.AppendLine("   <Session>");

                                                //Mandatory String
                                                sbSessionList.AppendLine("      <SessionId>" + varPalmClientDexId + "_" + varExt + "</SessionId>");

                                                //Mandatory String
                                                sbSessionList.AppendLine("      <CaseId>" + varPalmClientSupportId + "_" + varExt + "</CaseId>");

                                                //Mandatory Date
                                                sbSessionList.AppendLine("      <SessionDate>" + dbEntryDate + "</SessionDate>");

                                                //ToDo: Optional Integer
                                                sbSessionList.AppendLine("      <ServiceTypeId>" + varServTypeId + "</ServiceTypeId>");

                                                //Optional Integer
                                                if (varGetRecip > 0 && 1 == 2)
                                                    sbSessionList.AppendLine("      <TotalNumberOfUnidentifiedClients>" + varGetRecip + "</TotalNumberOfUnidentifiedClients>"); //Mandatory
                                                else
                                                    sbSessionList.AppendLine("      <TotalNumberOfUnidentifiedClients>0</TotalNumberOfUnidentifiedClients>"); //Mandatory

                                                // For CHSP only
                                                //if (varDoCHSP == true)
                                                //{
                                                //  sbSessionList.AppendLine("      <FeesCharged>0.00</FeesCharged>");
                                                //}

                                                //Optional True / False
                                                sbSessionList.AppendLine("      <InterpreterPresent>" + varInterpreterUsed + "</InterpreterPresent>");


                                                sbSessionList.AppendLine("      <SessionClients>");
                                                sbSessionList.AppendLine("         <SessionClient>");
                                                sbSessionList.AppendLine("            <ClientId>" + dbClient + "</ClientId>");
                                                //sbSessionList.AppendLine ("            <ParticipationTypeCode>CLIENT</ParticipationTypeCode>");
                                                sbSessionList.AppendLine("            <ParticipationCode>CLIENT</ParticipationCode>");
                                                sbSessionList.AppendLine("         </SessionClient>");
                                                sbSessionList.AppendLine("      </SessionClients>");

                                                // For CHSP only
                                                if (varDoCHSP == true)
                                                    sbSessionList.AppendLine("      <TimeMinutes>" + dbHours + "</TimeMinutes>");

                                                sbSessionList.AppendLine("   </Session>");

                                            } // varDoNext2

                                        } // Same support period

                                    } // Dex Loop

                                } // prevent dup case

                            } // client loop


                            //Header part of the DEX extract
                            sbHeaderList.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");

                            sbHeaderList.AppendLine("<DEXFileUpload>");

                            //Section 1: Clients
                            sbHeaderList.AppendLine("<Clients>");
                            sbHeaderList.AppendLine(sbClientList.ToString());
                            sbHeaderList.AppendLine("</Clients>");

                            //Section 2: Cases
                            sbHeaderList.AppendLine("<Cases>");
                            sbHeaderList.AppendLine(sbCaseList.ToString());
                            sbHeaderList.AppendLine("</Cases>");

                            //Section 3: Sessions
                            sbHeaderList.AppendLine("<Sessions>");
                            sbHeaderList.AppendLine(sbSessionList.ToString());
                            sbHeaderList.AppendLine("</Sessions>");

                            //Section 4: Session Assessments | Group Only
                            if (sbSessionAssessmentList.Length > 0)
                            {
                                sbHeaderList.AppendLine("<SessionAssessments>");
                                sbHeaderList.AppendLine(sbSessionAssessmentList.ToString());
                                sbHeaderList.AppendLine("</SessionAssessments>");
                            }

                            //Section 5: Client Assessments
                            sbHeaderList.AppendLine("<ClientAssessments>");
                            sbHeaderList.AppendLine(sbClientAssessmentList.ToString());
                            sbHeaderList.AppendLine("</ClientAssessments>");

                            sbHeaderList.AppendLine("</DEXFileUpload>");


                            //varTest += sbHeaderList.ToString();

                            // Create note against current Palm Go DEX record and add attachment
                            byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                            string encodedData = System.Convert.ToBase64String(filename);
                            Entity Annotation = new Entity("annotation");
                            Annotation.Attributes["objectid"] = new EntityReference("new_palmgodex", varDexID);
                            Annotation.Attributes["objecttypecode"] = "new_palmgodex";
                            Annotation.Attributes["subject"] = "DEX Extract";
                            Annotation.Attributes["documentbody"] = encodedData;
                            Annotation.Attributes["mimetype"] = @"text / plain";
                            Annotation.Attributes["notetext"] = "DEX Extract for " + cleanDate(varStartDatePr) + " to " + cleanDate(varEndDatePr);
                            Annotation.Attributes["filename"] = varFileName;
                            _service.Create(Annotation);

                            // If there is an error, create note against current Palm Go DEX record and add attachment
                            if (sbErrorList.Length > 0)
                            {
                                byte[] filename2 = Encoding.ASCII.GetBytes(sbErrorList.ToString());
                                string encodedData2 = System.Convert.ToBase64String(filename2);
                                Entity Annotation2 = new Entity("annotation");
                                Annotation2.Attributes["objectid"] = new EntityReference("new_palmgodex", varDexID);
                                Annotation2.Attributes["objecttypecode"] = "new_palmgodex";
                                Annotation2.Attributes["subject"] = "DEX Extract";
                                Annotation2.Attributes["documentbody"] = encodedData2;
                                Annotation2.Attributes["mimetype"] = @"text / plain";
                                Annotation2.Attributes["notetext"] = "DEX errors and warnings for " + cleanDate(varStartDatePr) + " to " + cleanDate(varEndDatePr);
                                Annotation2.Attributes["filename"] = varFileName2;
                                _service.Create(Annotation2);
                            }

                            //varTest += cleanDate(varStartDatePr) + " " + cleanDate(varEndDatePr);

                            //throw new InvalidPluginExecutionException("This plugin is working\r\n" + varTest);
                        } else if (varReport == ("External"))
                        {
                            // Fetch statements for database
                            // Get the required fields from the external records table (and associated entities)
                            // Any external records that have a DEX record for the period
                            string dbExternalERRecords = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_externalerrecord'>
                                <attribute name='new_externalerrecordid' />
                                <attribute name='new_externalidnum' />
                                <attribute name='new_firstname' />
                                <attribute name='new_lastname' />
                                <attribute name='new_dateofbirth' />
                                <attribute name='new_dateofbirthestimate' />
                                <attribute name='new_gender' />
                                <attribute name='new_countryofbirthid' />
                                <attribute name='new_languageid' />
                                <attribute name='new_indigenousstatus' />
                                <attribute name='new_yearofarrival' />
                                <attribute name='new_impairmentsconditionsdisabilities' />
                                <attribute name='new_suburbid' />
                                <attribute name='new_homeless' />
                                <attribute name='new_mainsourceofincome' />
                                <link-entity name='new_palmclientfinancial' from='new_palmclientfinancialid' to='new_financialid' link-type='outer'>
                                    <attribute name='new_palmclientfinancialid' />
                                    <attribute name='new_entrydate' />
                                    <attribute name='new_firstnamenc' />
                                    <attribute name='new_surnamenc' />
                                    <order attribute='new_entrydate' descending='true' />
                                    <link-entity name='new_palmclientdex' from='new_financial' to='new_palmclientfinancialid' link-type='inner'>
                                        <attribute name='new_entrydate' />
                                        <attribute name='new_agencyid' />
                                    </link-entity>
                                </link-entity>
                                <link-entity name='new_palmddllocality' from='new_palmddllocalityid' to='new_suburbid' link-type='inner'>
                                    <attribute name='new_postcode' />
                                    <attribute name='new_state' />
                                </link-entity>
                                <link-entity name='new_palmddlcountry' from='new_palmddlcountryid' to='new_countryofbirthid' link-type='outer'>
                                    <attribute name='new_code' />
                                </link-entity>
                                <link-entity name='new_palmddllanguage' from='new_palmddllanguageid' to='new_languageid' link-type='outer'>
                                    <attribute name='new_code' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname = 'new_palmclientdex' attribute = 'new_agencyid' operator= 'eq' value = '" + ownerLookup.Id + @"' />
                                </filter>
                              </entity>
                            </fetch> ";

                            // Get the required fields from the dex table
                            // Any DEX records against the agency id that fall within the period
                            string dbDEXList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientdex'>
                                <attribute name='new_entrydate' />
                                <attribute name='new_palmclientdexid' />
                                <attribute name='new_consentfuture' />
                                <attribute name='new_consentrelease' />
                                <attribute name='new_norecip' />
                                <attribute name='new_hours' />
                                <attribute name='new_servtype' />
                                <attribute name='new_financial' />
                                <attribute name='new_agencyid' />
                                <filter type='and'>
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientdex' attribute='new_agencyid' operator='eq' value='" + ownerLookup.Id + @"' />
                                </filter>
                                <order attribute='new_financial' />
                              </entity>
                            </fetch> ";

                            // Get the required fields from the dex scores table
                            // Any DEX scores records against the agency id that fall within the period
                            string dbScoresList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientdex'>
                                <attribute name='new_entrydate' />
                                <attribute name='new_palmclientdexid' />
                                <attribute name='new_consentfuture' />
                                <attribute name='new_consentrelease' />
                                <attribute name='new_norecip' />
                                <attribute name='new_hours' />
                                <attribute name='new_servtype' />
                                <attribute name='new_financial' />
                                <attribute name='new_agencyid' />
                                <link-entity name='new_palmclientdexscores' to='new_palmclientdexid' from='new_dexrecord' link-type='inner'>
                                    <attribute name='new_palmclientdexscoresid' />
                                    <attribute name='new_score' />
                                    <attribute name='new_section' />
                                    <attribute name='new_dexrecord' />
                                    <attribute name='new_thetype' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientdex' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientdex' attribute='new_agencyid' operator='eq' value='" + ownerLookup.Id + @"' />
                                </filter>
                                <order attribute='new_financial' />
                              </entity>
                            </fetch> ";

                            // Get the required fields from the SU Drop Down list entity
                            string dbDropDownList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                  <entity name='new_palmsudropdown'>
                                    <attribute name='new_type' />
                                    <attribute name='new_description' />
                                    <attribute name='new_er' />
                                    <order attribute='new_description' />
                                  </entity>
                                </fetch> ";

                            // Variables to hold database values
                            string dbExternalERRecordId = ""; 
                            string varExternalERRecordId = "";
                            string dbExternal = "";
                            string dbFirstName = "";
                            string dbSurname = "";
                            string dbDob = "";
                            string dbDobEst = "";
                            string dbGender = "";
                            string dbDexSlk = "";
                            string dbCountry = "";
                            string dbLanguage = "";
                            string dbIndigenous = "";
                            string dbYearArrival = "";
                            string dbPalmClientSupportId = "";
                            string varPalmClientSupportId = "";
                            string dbCCPDisType = "";
                            string varDisType = "";
                            string dbLocality = "";
                            string dbIsHomeless = "";
                            string dbLivingArrangePres = "";
                            string dbIncomePres = "";
                            string dbInterpreterUsed = "";
                            string dbSourceRef = "";
                            string dbPrimReason = "";
                            string dbReasons = "";
                            string dbResidentialPres = "";
                            string dbDva = "";
                            string dbCarerAvail = "";
                            string dbEntryDate = "";
                            string varEntryDate = "";
                            string varEntryDate2 = "";
                            string dbAgencyId = "";
                            string dbState = "";
                            string dbPostcode = "";

                            string dbEntryDate2 = "";
                            string dbPalmClientDexId = "";
                            string varPalmClientDexId = "";
                            string dbConsentFuture = "";
                            string dbConsentRelease = "";
                            string dbNoRecip = "";
                            string dbHours = "";
                            string dbServType = "";
                            string dbSupportPeriod = "";
                            string dbSupportPeriod2 = "";
                            string dbAgencyId2 = "";
                            string dbPalmClientDexScoresId = "";
                            string dbScore = "";
                            string dbSection = "";
                            string varSection = "";
                            string dbDexRecord = "";
                            string dbTheType = "";

                            // Drop down list variables
                            string varDesc = "";
                            string varType = "";
                            string varDEX = "";

                            // Variables for SLK
                            string varFirstName = "";
                            string varSurname = "";
                            string varGender = "";
                            string varDob = "";

                            // Variables for DEX record used to convert values
                            string varInterpreterUsed = "";
                            string varHasDis = "";
                            string varIsHomeless = "";

                            // Id variables
                            string varExternalNumber = "";
                            string varSessionNumber = "";
                            string varCaseNumber = "";

                            // Used to determine whether to process records
                            int varDoNext = 0;
                            int varDoNext2 = 0;
                            int varDoNext3 = 0;
                            int varMaxRecip = 0; // Maximum recipients

                            string varReasonStr = ""; // Reasons for presenting string

                            // Validate whether data was found
                            bool varSource = false;
                            bool varIndigenous = false;
                            bool varState = false;
                            bool varIncome = false;
                            bool varLiving = false;
                            bool varCCP = false;
                            bool varResidential = false;
                            bool varDV = false;

                            int varNoRecip = 0; // Number of recipients
                            int varGetRecip = 0; // Recipients from table

                            // Strings for adding DEX scores
                            string varPreGroup = "a";
                            string varPostGroup = "a";
                            string varPreCirc = "a";
                            string varPostCirc = "a";
                            string varPreGoal = "a";
                            string varPostGoal = "a";
                            string varSatisfaction = "a";
                            bool varSeeClient3 = false; // If client exists
                            string dbPalmClientDexId2 = ""; // Dex ID for comparison
                            string[] value; // Values string
                            string varServTypeId = ""; // Service type id

                            //Debug variable
                            string dbTemp = "";

                            // Get the fetch XML data and place in entity collection objects
                            EntityCollection result = _service.RetrieveMultiple(new FetchExpression(dbExternalERRecords));
                            EntityCollection result2 = _service.RetrieveMultiple(new FetchExpression(dbDEXList));
                            EntityCollection result3 = _service.RetrieveMultiple(new FetchExpression(dbDropDownList));
                            EntityCollection result8 = _service.RetrieveMultiple(new FetchExpression(dbScoresList));

                            // Strings to check for duplicates
                            varExternalNumber = "...";
                            varSessionNumber = "...";
                            varCaseNumber = "...";

                            // Loop through client records
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

                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (c.FormattedValues.Contains("new_externalerrecordid"))
                                    dbExternalERRecordId = c.FormattedValues["new_externalerrecordid"];
                                else if (c.Attributes.Contains("new_externalerrecordid"))
                                    dbExternalERRecordId = c.Attributes["new_externalerrecordid"].ToString();
                                else
                                    dbExternalERRecordId = "";

                                // Remove dashes
                                varExternalERRecordId = "";
                                if (String.IsNullOrEmpty(dbExternalERRecordId) == false)
                                    varExternalERRecordId = dbExternalERRecordId.Replace("-", "");


                                //External Id num
                                if (c.FormattedValues.Contains("new_externalidnum"))
                                    dbExternal = c.FormattedValues["new_externalidnum"];
                                else if (c.Attributes.Contains("new_externalidnum"))
                                    dbExternal = c.Attributes["new_externalidnum"].ToString();
                                else
                                    dbExternal = "";

                                //Reset duplicate client checker
                                varDoNext = 0;

                                if (varExternalNumber.IndexOf("*" + dbExternalERRecordId + "*") > -1)
                                {                            
                                    varDoNext = 1;
                                }

                                //varExternalERRecordId += "*" + dbExternalERRecordId + "*";
                                varExternalNumber += "*" + dbExternalERRecordId + "*";

                                if (c.FormattedValues.Contains("new_firstname"))
                                    dbFirstName = c.FormattedValues["new_firstname"];
                                else if (c.Attributes.Contains("new_firstname"))
                                    dbFirstName = c.Attributes["new_firstname"].ToString();
                                else
                                    dbFirstName = "";

                                if (c.FormattedValues.Contains("new_lastname"))
                                    dbSurname = c.FormattedValues["new_lastname"];
                                else if (c.Attributes.Contains("new_lastname"))
                                    dbSurname = c.Attributes["new_lastname"].ToString();
                                else
                                    dbSurname = "";

                                if (c.FormattedValues.Contains("new_dateofbirth"))
                                    dbDob = c.FormattedValues["new_dateofbirth"];
                                else if (c.Attributes.Contains("new_dateofbirth"))
                                    dbDob = c.Attributes["new_dateofbirth"].ToString();
                                else
                                    dbDob = "";

                                // Convert date from American format to Australian format
                                dbDob = cleanDateAM(dbDob);

                                if (c.FormattedValues.Contains("new_dateofbirthestimate"))
                                    dbDobEst = c.FormattedValues["new_dateofbirthestimate"];
                                else if (c.Attributes.Contains("new_dateofbirthestimate"))
                                    dbDobEst = c.Attributes["new_dateofbirthestimate"].ToString();
                                else
                                    dbDobEst = "";

                                if (dbDobEst == "Yes" || dbDobEst == "true")
                                    dbDobEst = "true";
                                else
                                    dbDobEst = "false";

                                // Convert DOB to valid format
                                if (DateTime.TryParse(dbDob, out varCheckDate))
                                    dbDob = cleanDateE(varCheckDate);
                                else
                                {
                                    varCheckDate = Convert.ToDateTime("1-Jan-1970");
                                    dbDob = "1970-01-01";
                                }

                                if (c.FormattedValues.Contains("new_gender"))
                                    dbGender = c.FormattedValues["new_gender"];
                                else if (c.Attributes.Contains("new_gender"))
                                    dbGender = c.Attributes["new_gender"].ToString();
                                else
                                    dbGender = "";

                                // Create SLK based on firstname, surname, gender and dob
                                varSurname = dbSurname.Replace(" ", string.Empty); //Prepare surname by removing spaces.
                                                                                   //Need to clean surname.
                                varSurname = cleanString(varSurname.ToUpper(), "slk");
                                varSurname = varSurname + "22222";
                                varSurname = varSurname.Substring(1, 1) + varSurname.Substring(2, 1) + varSurname.Substring(4, 1);

                                if (varSurname == "222")
                                    varSurname = "999";

                                //Get the gender code
                                if (dbGender == "Female")
                                    varGender = "2";
                                else if (dbGender == "Male")
                                    varGender = "1";
                                else
                                    varGender = "9";

                                //Put dob into expected format
                                varDob = cleanDateS(varCheckDate);

                                //Get the statistical linkage key
                                dbDexSlk = varSurname + varFirstName + varDob + varGender;


                                dbDexSlk = dbDexSlk.ToUpper();
                                if (dbDexSlk.Length != 14)
                                    sbErrorList.AppendLine("(" + dbExternal + ") " + dbFirstName + " " + dbSurname + ": SLK not 14 characters long | " + dbDexSlk + "<br>");

                                if (c.FormattedValues.Contains("new_palmddlcountry4.new_code"))
                                    dbCountry = c.FormattedValues["new_palmddlcountry4.new_code"];
                                else if (c.Attributes.Contains("new_palmddlcountry4.new_code"))
                                    dbCountry = c.Attributes["new_palmddlcountry4.new_code"].ToString();
                                else
                                    dbCountry = "";

                                if (c.FormattedValues.Contains("new_palmddllanguage5.new_code"))
                                    dbLanguage = c.FormattedValues["new_palmddllanguage5.new_code"];
                                else if (c.Attributes.Contains("new_palmddllanguage5.new_code"))
                                    dbLanguage = c.Attributes["new_palmddllanguage5.new_code"].ToString();
                                else
                                    dbLanguage = "";

                                //varTest += dbCountry + " " + dbLanguage + "\r\n";

                                if (c.FormattedValues.Contains("new_indigenousstatus"))
                                    dbIndigenous = c.FormattedValues["new_indigenousstatus"];
                                else if (c.Attributes.Contains("new_indigenousstatus"))
                                    dbIndigenous = c.Attributes["new_indigenousstatus"].ToString();
                                else
                                    dbIndigenous = "";

                                if (c.FormattedValues.Contains("new_yearofarrival"))
                                    dbYearArrival = c.FormattedValues["new_yearofarrival"];
                                else if (c.Attributes.Contains("new_yearofarrival"))
                                    dbYearArrival = c.Attributes["new_yearofarrival"].ToString();
                                else
                                    dbYearArrival = "";

                                // Ensure numeric
                                dbYearArrival = cleanString(dbYearArrival, "number");

                                if (c.FormattedValues.Contains("new_palmclientfinancial1.new_palmclientfinancialid"))
                                    dbPalmClientSupportId = c.FormattedValues["new_palmclientfinancial1.new_palmclientfinancialid"];
                                else if (c.Attributes.Contains("new_palmclientfinancial1.new_palmclientfinancialid"))
                                    dbPalmClientSupportId = c.GetAttributeValue<AliasedValue>("new_palmclientfinancial1.new_palmclientfinancialid").Value.ToString();
                                else
                                    dbPalmClientSupportId = "";

                                // Remove dashes
                                varPalmClientSupportId = "";
                                if (String.IsNullOrEmpty(dbPalmClientSupportId) == false)
                                    varPalmClientSupportId = dbPalmClientSupportId.Replace("-", "");

                                if (c.FormattedValues.Contains("new_impairmentsconditionsdisabilities"))
                                    dbCCPDisType = c.FormattedValues["new_impairmentsconditionsdisabilities"];
                                else if (c.Attributes.Contains("new_impairmentsconditionsdisabilities"))
                                    dbCCPDisType = c.Attributes["new_impairmentsconditionsdisabilities"].ToString();
                                else
                                    dbCCPDisType = "";

                                dbCCPDisType = getMult(dbCCPDisType); //External ER - disability is multiple options.

                                if (c.FormattedValues.Contains("new_suburbid"))
                                    dbLocality = c.FormattedValues["new_suburbid"];
                                else if (c.Attributes.Contains("new_suburbid"))
                                    dbLocality = c.Attributes["new_suburbid"].ToString();
                                else
                                    dbLocality = "";

                                if (c.FormattedValues.Contains("new_palmddllocality3.new_state"))
                                    dbState = c.FormattedValues["new_palmddllocality3.new_state"];
                                else if (c.Attributes.Contains("new_palmddllocality3.new_state"))
                                    dbState = c.Attributes["new_palmddllocality3.new_state"].ToString();
                                else
                                    dbState = "";

                                if (c.FormattedValues.Contains("new_palmddllocality3.new_postcode"))
                                    dbPostcode = c.FormattedValues["new_palmddllocality3.new_postcode"];
                                else if (c.Attributes.Contains("new_palmddllocality3.new_postcode"))
                                    dbPostcode = c.Attributes["new_palmddllocality3.new_postcode"].ToString();
                                else
                                    dbPostcode = "";

                                if (c.FormattedValues.Contains("new_homeless"))
                                    dbIsHomeless = c.FormattedValues["new_homeless"];
                                else if (c.Attributes.Contains("new_homeless"))
                                    dbIsHomeless = c.Attributes["new_homeless"].ToString();
                                else
                                    dbIsHomeless = "";

                                // Wrap asterisks around values for better comparisons
                                dbIsHomeless = getMult(dbIsHomeless);



                                //DON'T KNOW ABOUT THESE

                                //Conditional Mandatory
                                /*
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_livingarrangepres"))
                                    dbLivingArrangePres = c.FormattedValues["new_palmclientsupport1.new_livingarrangepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_livingarrangepres"))
                                    dbLivingArrangePres = c.Attributes["new_palmclientsupport1.new_livingarrangepres"].ToString();
                                else
                                    dbLivingArrangePres = "";
                                */

                                //Optional
                                /*
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_interpreterused"))
                                    dbInterpreterUsed = c.FormattedValues["new_palmclientsupport1.new_interpreterused"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_interpreterused"))
                                    dbInterpreterUsed = c.Attributes["new_palmclientsupport1.new_interpreterused"].ToString();
                                else
                                    dbInterpreterUsed = "";
                                */

                                //optional
                                /*
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_sourceref"))
                                    dbSourceRef = c.FormattedValues["new_palmclientsupport1.new_sourceref"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_sourceref"))
                                    dbSourceRef = c.Attributes["new_palmclientsupport1.new_sourceref"].ToString();
                                else
                                    dbSourceRef = "";
                                */

                                //optional
                                /*
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_primreason"))
                                    dbPrimReason = c.FormattedValues["new_palmclientsupport1.new_primreason"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_primreason"))
                                    dbPrimReason = c.Attributes["new_palmclientsupport1.new_primreason"].ToString();
                                else
                                    dbPrimReason = "";
                                */

                                //optional
                                /*
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_reasons"))
                                    dbReasons = c.FormattedValues["new_palmclientsupport1.new_reasons"]; // multi
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_reasons"))
                                    dbReasons = c.Attributes["new_palmclientsupport1.new_reasons"].ToString();
                                else
                                    dbReasons = "";

                                dbReasons = getMult(dbReasons);
                                */

                                //conditional mandatory
                                /*
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_residentialpres"))
                                    dbResidentialPres = c.FormattedValues["new_palmclientsupport1.new_residentialpres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_residentialpres"))
                                    dbResidentialPres = c.Attributes["new_palmclientsupport1.new_residentialpres"].ToString();
                                else
                                    dbResidentialPres = "";
                                */

                                //conditional mandatory
                                /*
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_dva"))
                                    dbDva = c.FormattedValues["new_palmclientsupport1.new_dva"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_dva"))
                                    dbDva = c.Attributes["new_palmclientsupport1.new_dva"].ToString();
                                else
                                    dbDva = "";
                                */

                                //optional
                                /*
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_careravail"))
                                    dbCarerAvail = c.FormattedValues["new_palmclientsupport1.new_careravail"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_careravail"))
                                    dbCarerAvail = c.Attributes["new_palmclientsupport1.new_careravail"].ToString();
                                else
                                    dbCarerAvail = "";
                                */
                                //UNSURE ON ABOVE


                                if (c.FormattedValues.Contains("new_mainsourceofincome"))
                                    dbIncomePres = c.FormattedValues["new_mainsourceofincome"];
                                else if (c.Attributes.Contains("new_mainsourceofincome"))
                                    dbIncomePres = c.Attributes["new_mainsourceofincome"].ToString();
                                else
                                    dbIncomePres = "";

                                if (c.FormattedValues.Contains("new_palmclientdex2.new_agencyid"))
                                    dbAgencyId = c.FormattedValues["new_palmclientdex2.new_agencyid"];
                                else if (c.Attributes.Contains("new_palmclientdex2.new_agencyid"))
                                    dbAgencyId = c.GetAttributeValue<AliasedValue>("new_palmclientdex2.new_agencyid").Value.ToString();
                                else
                                    dbAgencyId = "";

                                if (c.FormattedValues.Contains("new_palmclientdex2.new_entrydate"))
                                    dbEntryDate = c.FormattedValues["new_palmclientdex2.new_entrydate"];
                                else if (c.Attributes.Contains("new_palmclientdex2.new_entrydate"))
                                    dbEntryDate = c.Attributes["new_palmclientdex2.new_entrydate"].ToString();
                                else
                                    dbEntryDate = "";

                                // Convert date from American format to Australian format
                                dbEntryDate = cleanDateAM(dbEntryDate);

                                if (String.IsNullOrEmpty(dbEntryDate) == false)
                                    varEntryDate = dbEntryDate.Replace("-", "_");
                                else
                                    varEntryDate = "-";

                                // Convert date to required format
                                if (String.IsNullOrEmpty(dbEntryDate) == false)
                                    dbEntryDate = cleanDateE(Convert.ToDateTime(dbEntryDate));

                                // Get required values based on data
                                if (dbInterpreterUsed == "Yes" || dbInterpreterUsed == "yes")
                                    varInterpreterUsed = "true";
                                else
                                    varInterpreterUsed = "false";
                                
                                if(String.IsNullOrEmpty(dbCCPDisType) == false && dbCCPDisType != "-" && dbCCPDisType.ToLower().Contains("none") == false && dbCCPDisType.ToLower().Contains("not stated/inadequately described") == false)
                                    varHasDis = "true";
                                else
                                    varHasDis = "false";

                                if (dbCCPDisType.ToLower().Contains("intellectual learning"))
                                    varDisType = "LEARNING";
                                else if (dbCCPDisType.ToLower().Contains("physical/diverse"))
                                    varDisType = "PHYSICAL";
                                else if (dbCCPDisType.ToLower().Contains("psychiatric"))
                                    varDisType = "PSYCHIATRIC";
                                else if (dbCCPDisType.ToLower().Contains("sensory/speech"))
                                    varDisType = "SENSORY";

                                if (dbIsHomeless.ToLower().IndexOf("short-term or emergency accom") > -1 || dbIsHomeless.ToLower().IndexOf("sleeping rough") > -1)
                                    varIsHomeless = "Yes";
                                else
                                    varIsHomeless = "No";

                                // Ensure postcode is 4 digits
                                if (dbPostcode.Length == 1)
                                    dbPostcode = "000" + dbPostcode;
                                if (dbPostcode.Length == 2)
                                    dbPostcode = "00" + dbPostcode;
                                if (dbPostcode.Length == 3)
                                    dbPostcode = "0" + dbPostcode;

                                varMaxRecip = 0; // Reset

                                // Only do this if client has not already been processed
                                if (varDoNext == 0)
                                {
                                    //Data Validation
                                    if (String.IsNullOrEmpty(dbLocality) == true)
                                        sbErrorList.AppendLine("(" + dbExternal + ") " + dbFirstName + " " + dbSurname + ": Missing Locality<br>");

                                    if (String.IsNullOrEmpty(dbState) == true)
                                        sbErrorList.AppendLine("(" + dbExternal + ") " + dbFirstName + " " + dbSurname + ": Missing State<br>");

                                    if (String.IsNullOrEmpty(dbPostcode) == true)
                                        sbErrorList.AppendLine("(" + dbExternal + ") " + dbFirstName + " " + dbSurname + ": Missing Postcode<br>");

                                    //if (String.IsNullOrEmpty(dbLivingArrangePres) == true)
                                      //  sbErrorList.AppendLine("(" + dbExternal + ") " + dbFirstName + " " + dbSurname + ": Missing Living Arrangements<br>");

                                    if (String.IsNullOrEmpty(dbIncomePres) == true)
                                        sbErrorList.AppendLine("(" + dbExternal + ") " + dbFirstName + " " + dbSurname + ": Missing Income<br>");

                                   // if (String.IsNullOrEmpty(dbSourceRef) == true)
                                      //  sbErrorList.AppendLine("(" + dbExternal + ") " + dbFirstName + " " + dbSurname + ": Missing Referral Source<br>");

                                    //if (String.IsNullOrEmpty(dbPrimReason) == true)
                                       // sbErrorList.AppendLine("(" + dbExternal + ") " + dbFirstName + " " + dbSurname + ": Missing Primary Presenting Reason<br>");

                                    dbGender = dbGender.ToUpper();
                                    if (dbGender != "MALE" && dbGender != "FEMALE")
                                        dbGender = "NOTSTATED";

                                    // Reset data checkers
                                    varIndigenous = false;
                                    varState = false;
                                    varIncome = false;
                                    varLiving = false;
                                    varSource = false;
                                    varCCP = false;
                                    varResidential = false;
                                    varDV = false;

                                    if (dbCountry == "Don't Know" || dbCountry == "Not Applicable")
                                        dbCountry = "9999";

                                    // Loop through drop down list values
                                    foreach (var d in result3.Entities)
                                    {
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

                                        if (d.FormattedValues.Contains("new_er"))
                                            varDEX = d.FormattedValues["new_er"];
                                        else if (d.Attributes.Contains("new_er"))
                                            varDEX = d.Attributes["new_er"].ToString();
                                        else
                                            varDEX = "";

                                        /*
                                        // Get DEX value based on type and whether the description matches the data
                                        if (varType.ToLower() == "sourceref" && dbSourceRef.ToLower() == varDesc.ToLower())
                                        {
                                            dbSourceRef = varDEX;
                                            varSource = true;
                                        }*/


                                        if (varType.ToLower() == "indigenous" && dbIndigenous.ToLower() == varDesc.ToLower())
                                        {
                                            dbIndigenous = varDEX;
                                            varIndigenous = true;
                                        }

                                        if (varType.ToLower() == "state" && dbState.ToLower() == varDesc.ToLower())
                                        {
                                            dbState = varDEX;
                                            varState = true;
                                        }

                                        if (varType.ToLower() == "income" && dbIncomePres.ToLower() == varDesc.ToLower())
                                        {
                                            dbIncomePres = varDEX;
                                            varIncome = true;
                                        }

                                        /*
                                        if (varType.ToLower() == "livingarrange" && dbLivingArrangePres.ToLower() == varDesc.ToLower())
                                        {
                                            dbLivingArrangePres = varDEX;
                                            varLiving = true;
                                        }*/

                                        /*
                                        if (varType.ToLower() == "ccpdis" && dbCCPDisType.ToLower() == varDesc.ToLower())
                                        {
                                            dbCCPDisType = varDEX;
                                            varCCP = true;
                                        }*/

                                        /*
                                        if (varType.ToLower() == "residential" && dbResidentialPres.ToLower() == varDesc.ToLower())
                                        {
                                            dbResidentialPres = varDEX;
                                            varResidential = true;
                                        }*/

                                        /*
                                        if (varType.ToLower() == "dva" && dbDva.ToLower() == varDesc.ToLower())
                                        {
                                            dbDva = varDEX;
                                            varDV = true;
                                        }*/

                                    } // Drop List

                                    // Set defaults for missing data
                                    if (varSource == false)
                                        dbSourceRef = "NOTSTATED";

                                    if (varIndigenous == false)
                                        dbIndigenous = "NOTSTATED";

                                    if (varState == false)
                                        dbState = "VIC";

                                    if (varIncome == false)
                                        dbIncomePres = "NOTSTATED";

                                    if (varLiving == false)
                                        dbLivingArrangePres = "NOTSTATED";

                                    if (varCCP == false)
                                        dbCCPDisType = "NOTSTATED";

                                    if (varResidential == false)
                                        dbResidentialPres = "NOTSTATED";

                                    if (varDV == false)
                                        dbDva = "NODVA";

                                    // Convert to valid values
                                    if (dbConsentFuture == "Yes" || dbConsentFuture == "true")
                                        dbConsentFuture = "true";
                                    else
                                        dbConsentFuture = "false";

                                    if (dbConsentRelease == "Yes" || dbConsentRelease == "true")
                                        dbConsentRelease = "true";
                                    else
                                        dbConsentRelease = "false";


                                    // Create client part of extract
                                    sbClientList.AppendLine("   <Client>");

                                    //Mandatory String
                                    sbClientList.AppendLine("      <ClientId>" + dbExternal + "</ClientId>");

                                    //Optional String - Mandatory if consent false
                                    if (dbDexSlk != "-")
                                        sbClientList.AppendLine("      <Slk>" + dbDexSlk + "</Slk>");

                                    if (dbDexSlk == "-" && dbConsentRelease == "false")
                                        sbErrorList.AppendLine("(" + dbExternal + ") " + dbFirstName + " " + dbSurname + " has not given consent and does not have an SLK<br>");

                                    //Mandatory True / False
                                    if (String.IsNullOrEmpty(dbConsentRelease) == false)
                                        sbClientList.AppendLine("      <ConsentToProvideDetails>" + dbConsentRelease + "</ConsentToProvideDetails>");
                                    else
                                        sbClientList.AppendLine("      <ConsentToProvideDetails>false</ConsentToProvideDetails>");

                                    if (String.IsNullOrEmpty(dbConsentFuture) == false)
                                        sbClientList.AppendLine("      <ConsentedForFutureContacts>" + dbConsentFuture + "</ConsentedForFutureContacts>");
                                    else
                                        sbClientList.AppendLine("      <ConsentedForFutureContacts>false</ConsentedForFutureContacts>");

                                    //Conditional String - Not allowed if Consent False
                                    if (dbConsentRelease == "true")
                                    {
                                        sbClientList.AppendLine("      <GivenName>" + dbFirstName + "</GivenName>");
                                        sbClientList.AppendLine("      <FamilyName>" + dbSurname + "</FamilyName>");
                                    }

                                    //Always False
                                    sbClientList.AppendLine("      <IsUsingPsuedonym>false</IsUsingPsuedonym>");

                                    //Mandatory YYYY-MM-DD
                                    sbClientList.AppendLine("      <BirthDate>" + dbDob + "</BirthDate>");
                                    //Mandatory True / False (if true DOB must be YYYY-01-01)
                                    sbClientList.AppendLine("      <IsBirthDateAnEstimate>" + dbDobEst + "</IsBirthDateAnEstimate>");

                                    //Mandatory Reference Data
                                    sbClientList.AppendLine("      <GenderCode>" + dbGender + "</GenderCode>");
                                    sbClientList.AppendLine("      <CountryOfBirthCode>" + dbCountry + "</CountryOfBirthCode>");
                                    sbClientList.AppendLine("      <LanguageSpokenAtHomeCode>" + dbLanguage + "</LanguageSpokenAtHomeCode>");
                                    sbClientList.AppendLine("      <AboriginalOrTorresStraitIslanderOriginCode>" + dbIndigenous + "</AboriginalOrTorresStraitIslanderOriginCode>");

                                    sbClientList.AppendLine("      <HasDisabilities>" + varHasDis + "</HasDisabilities>"); //Mandatory True / False

                                    //XML if has disabilities true
                                    if (varHasDis == "true")
                                    {
                                        sbClientList.AppendLine("      <Disabilities>");
                                        sbClientList.AppendLine("         <DisabilityCode>" + dbCCPDisType + "</DisabilityCode>"); //Mandatory Reference
                                        sbClientList.AppendLine("      </Disabilities>");
                                    }


                                    //if (varDoCHSP == true)
                                    //{
                                    sbClientList.AppendLine("      <AccommodationTypeCode>" + dbResidentialPres + "</AccommodationTypeCode>");
                                    sbClientList.AppendLine("      <DVACardStatusCode>" + dbDva + "</DVACardStatusCode>");

                                    /*
                                    if (dbCarerAvail == "Yes" || dbCarerAvail == "Has A Carer")
                                        sbClientList.AppendLine("      <HasCarer>true</HasCarer>");
                                    else
                                        sbClientList.AppendLine("      <HasCarer>false</HasCarer>");
                                    //}*/


                                    //XML for address
                                    sbClientList.AppendLine("      <ResidentialAddress>");

                                    //Optional String - not using
                                    //sbClientList.AppendLine("         <AddressLine1>[...]</AddressLine1>");
                                    //sbClientList.AppendLine("         <AddressLine2>[...]</AddressLine2>");

                                    //Mandatory String, reference and 4 character string
                                    sbClientList.AppendLine("         <Suburb>" + dbLocality + "</Suburb>");
                                    sbClientList.AppendLine("         <StateCode>" + dbState + "</StateCode>");
                                    sbClientList.AppendLine("         <Postcode>" + dbPostcode + "</Postcode>");
                                    sbClientList.AppendLine("      </ResidentialAddress>");

                                    //Optional True / False
                                    sbClientList.AppendLine("      <HomelessIndicatorCode>" + varIsHomeless + "</HomelessIndicatorCode>");

                                    //Optional Reference
                                    sbClientList.AppendLine("      <HouseholdCompositionCode>" + dbLivingArrangePres + "</HouseholdCompositionCode>");
                                    //Mandatory Reference
                                    sbClientList.AppendLine("      <MainSourceOfIncomeCode>" + dbIncomePres + "</MainSourceOfIncomeCode>");

                                    //sbClientList.AppendLine("      <IncomeFrequencyCode>[...]</IncomeFrequencyCode>"); //Optional Reference - not using
                                    //sbClientList.AppendLine("      <IncomeAmount>[...]</IncomeAmount>"); //Optional Number - not using

                                    // Add year of arrival
                                    if (dbYearArrival != "-" && dbYearArrival != "0" && String.IsNullOrEmpty(dbYearArrival) == false)
                                    {
                                        sbClientList.AppendLine("      <FirstArrivalYear>" + dbYearArrival + "</FirstArrivalYear>"); //Optional Number - greater/equal DOB, less than Today
                                        sbClientList.AppendLine("      <FirstArrivalMonth>January</FirstArrivalMonth>"); //Mandatory - January - "December led to issues with clients that arrived earlier in the same year."
                                    }

                                    //sbClientList.AppendLine("      <MigrationVisaCategoryCode>[...]</MigrationVisaCategoryCode>"); //Optional Reference - not using
                                    //sbClientList.AppendLine("      <AncestryCode>[...]</AncestryCode>"); //Optional Reference - not using

                                    sbClientList.AppendLine("   </Client>");

                                } // Client already processed


                                //Reset duplicate support period test
                                varDoNext3 = 0;
                                if (varCaseNumber.IndexOf("*" + dbPalmClientSupportId + "*") > -1)
                                    varDoNext3 = 1;

                                varCaseNumber += "*" + dbPalmClientSupportId + "*";

                                // Only do if support period not already processed
                                if (varDoNext3 == 0)
                                {

                                    // Build Case Here
                                    sbCaseList.AppendLine("   <Case>");

                                    //Mandatory String
                                    sbCaseList.AppendLine("      <CaseId>" + varPalmClientSupportId + "_" + varExt + "</CaseId>");

                                    //ToDo: Mandatory Integer
                                    sbCaseList.AppendLine("      <OutletActivityId>" + varOutletActivityId + "</OutletActivityId>");

                                    //Mandatory Integer
                                    if (varMaxRecip > 0)
                                        sbCaseList.AppendLine("      <TotalNumberOfUnidentifiedClients>" + varMaxRecip + "</TotalNumberOfUnidentifiedClients>");
                                    else
                                        sbCaseList.AppendLine("      <TotalNumberOfUnidentifiedClients>0</TotalNumberOfUnidentifiedClients>");

                                    //Optional
                                    sbCaseList.AppendLine("      <CaseClients>");

                                    sbCaseList.AppendLine("         <CaseClient>");

                                    sbCaseList.AppendLine("            <ClientId>" + dbExternal + "</ClientId>");
                                    //sbCaseList.AppendLine("            <ReferralSourceCode>" + dbSourceRef.ToUpper() + "</ReferralSourceCode>");



                                    // Append reasons for assistance if reasons not empty
                                    if (String.IsNullOrEmpty(dbReasons) == false)
                                    {

                                        sbCaseList.AppendLine("            <ReasonsForAssistance>");
                                        varReasonStr = "";
                                        // Loop through drop down list and get integer for primary reason
                                        foreach (var d in result3.Entities)
                                        {

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

                                            if (d.FormattedValues.Contains("new_er"))
                                                varDEX = d.FormattedValues["new_er"];
                                            else if (d.Attributes.Contains("new_er"))
                                                varDEX = d.Attributes["new_er"].ToString();
                                            else
                                                varDEX = "";

                                            // Add reason
                                            if (varDesc.ToLower() == dbPrimReason.ToLower() && varType == "reasons")
                                            {
                                                sbCaseList.AppendLine("               <ReasonForAssistance>");
                                                sbCaseList.AppendLine("                  <ReasonForAssistanceCode>" + varDEX + "</ReasonForAssistanceCode>");
                                                sbCaseList.AppendLine("                  <IsPrimary>true</IsPrimary>");
                                                sbCaseList.AppendLine("               </ReasonForAssistance>");

                                                varReasonStr += "*" + varDEX + "*";

                                                break;
                                            }

                                        } //k Loop


                                        // Loop through drop down list and get reasons for presenting
                                        foreach (var d in result3.Entities)
                                        {

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

                                            if (d.FormattedValues.Contains("new_er"))
                                                varDEX = d.FormattedValues["new_er"];
                                            else if (d.Attributes.Contains("new_er"))
                                                varDEX = d.Attributes["new_er"].ToString();
                                            else
                                                varDEX = "";

                                            // Add reason
                                            varTest = dbReasons;
                                            //sbCaseList.AppendLine("<DebugSection1>Greater than -1? " + dbReasons.IndexOf("*" + varDesc + "*") + " .. Not equal to " + dbPrimReason.ToLower() + "? " + varDesc.ToLower() + " .. Equal to -1?" + varReasonStr.IndexOf("*" + varDEX + "*") + " .. reasons?" + varType);
                                            if (dbReasons.IndexOf("*" + varDesc + "*") > -1 && varDesc.ToLower() != dbPrimReason.ToLower() && varReasonStr.IndexOf("*" + varDEX + "*") == -1 && varType == "reasons")
                                            {
                                                sbCaseList.AppendLine("               <ReasonForAssistance>");
                                                sbCaseList.AppendLine("                  <ReasonForAssistanceCode>" + varDEX + "</ReasonForAssistanceCode>");
                                                sbCaseList.AppendLine("                  <IsPrimary>false</IsPrimary>");
                                                sbCaseList.AppendLine("               </ReasonForAssistance>");

                                                varReasonStr += "*" + varDEX + "*";
                                            }

                                        } //j Loop

                                        sbCaseList.AppendLine("            </ReasonsForAssistance>");

                                    } //Reasons

                                    sbCaseList.AppendLine("         </CaseClient>");

                                    sbCaseList.AppendLine("      </CaseClients>");

                                    sbCaseList.AppendLine("   </Case>");

                                    //if (dbClient == "PLM31076")
                                    //    varTest += "Support Id: " + dbPalmClientSupportId + "\r\n";

                                    // Loop through data from DEX table
                                    foreach (var x in result2.Entities)
                                    {
                                        //varTest = "STARTING ATTRIBUTES:\r\n";

                                        //foreach (KeyValuePair<String, Object> attribute in s.Attributes)
                                        //{
                                        //    varTest += (attribute.Key + ": " + attribute.Value + "\r\n");
                                        //}

                                        //varTest += "STARTING FORMATTED:\r\n";

                                        //foreach (KeyValuePair<String, String> value in s.FormattedValues)
                                        //{
                                        //    varTest += (value.Key + ": " + value.Value + "\r\n");
                                        //}

                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (x.FormattedValues.Contains("new_palmclientdexid"))
                                            dbPalmClientDexId = x.FormattedValues["new_palmclientdexid"];
                                        else if (x.Attributes.Contains("new_palmclientdexid"))
                                            dbPalmClientDexId = x.Attributes["new_palmclientdexid"].ToString();
                                        else
                                            dbPalmClientDexId = "";

                                        // Remove dashes
                                        varPalmClientDexId = "";
                                        if (String.IsNullOrEmpty(dbPalmClientDexId) == false)
                                            varPalmClientDexId = dbPalmClientDexId.Replace("-", "");

                                        if (x.FormattedValues.Contains("new_entrydate"))
                                            dbEntryDate2 = x.FormattedValues["new_entrydate"];
                                        else if (x.Attributes.Contains("new_entrydate"))
                                            dbEntryDate2 = x.Attributes["new_entrydate"].ToString();
                                        else
                                            dbEntryDate2 = "";

                                        // Convert date from American format to Australian format
                                        dbEntryDate2 = cleanDateAM(dbEntryDate2);

                                        if (String.IsNullOrEmpty(dbEntryDate2) == false)
                                            varEntryDate2 = dbEntryDate2.Replace("-", "_");
                                        else
                                            varEntryDate2 = "-";

                                        if (x.FormattedValues.Contains("new_consentfuture"))
                                            dbConsentFuture = x.FormattedValues["new_consentfuture"];
                                        else if (x.Attributes.Contains("new_consentfuture"))
                                            dbConsentFuture = x.Attributes["new_consentfuture"].ToString();
                                        else
                                            dbConsentFuture = "";

                                        if (x.FormattedValues.Contains("new_consentrelease"))
                                            dbConsentRelease = x.FormattedValues["new_consentrelease"];
                                        else if (x.Attributes.Contains("new_consentrelease"))
                                            dbConsentRelease = x.Attributes["new_consentfuture"].ToString();
                                        else
                                            dbConsentRelease = "";

                                        if (x.FormattedValues.Contains("new_norecip"))
                                            dbNoRecip = x.FormattedValues["new_norecip"];
                                        else if (x.Attributes.Contains("new_norecip"))
                                            dbNoRecip = x.Attributes["new_norecip"].ToString();
                                        else
                                            dbNoRecip = "";

                                        dbNoRecip = cleanString(dbNoRecip, "number");
                                        Int32.TryParse(dbNoRecip, out varCheckInt);

                                        varNoRecip = varCheckInt;
                                        if (varNoRecip > varGetRecip)
                                            varGetRecip = varNoRecip;

                                        if (x.FormattedValues.Contains("new_hours"))
                                            dbHours = x.FormattedValues["new_hours"];
                                        else if (x.Attributes.Contains("new_hours"))
                                            dbHours = x.Attributes["new_hours"].ToString();
                                        else
                                            dbHours = "";

                                        dbHours = cleanString(dbHours, "double");

                                        Double.TryParse(dbHours, out varCheckDouble);
                                        dbHours = (int)(varCheckDouble * 60) + "";

                                        if (x.FormattedValues.Contains("new_servtype"))
                                            dbServType = x.FormattedValues["new_servtype"];
                                        else if (x.Attributes.Contains("new_servtype"))
                                            dbServType = x.Attributes["new_servtype"].ToString();
                                        else
                                            dbServType = "";

                                        // We need to get the entity id for the financial field for comparisons
                                        if (x.Attributes.Contains("new_financial"))
                                        {
                                            // Get the entity id for the client using the entity reference object
                                            getEntity = (EntityReference)x.Attributes["new_financial"];
                                            dbSupportPeriod = getEntity.Id.ToString();
                                        }
                                        else if (x.FormattedValues.Contains("new_financial"))
                                            dbSupportPeriod = x.FormattedValues["new_financial"];
                                        else
                                            dbSupportPeriod = "";

                                        // We need to get the entity id for the agency id field for comparisons
                                        if (x.Attributes.Contains("new_agencyid"))
                                        {
                                            // Get the entity id for the client using the entity reference object
                                            getEntity = (EntityReference)x.Attributes["new_agencyid"];
                                            dbAgencyId2 = getEntity.Id.ToString();
                                        }
                                        else if (x.FormattedValues.Contains("new_agencyid"))
                                            dbAgencyId2 = x.FormattedValues["new_agencyid"];
                                        else
                                            dbAgencyId2 = "";

                                        // Only do if the financials match
                                        if (dbSupportPeriod == dbPalmClientSupportId)
                                        {
                                            // Check to see if session has already been processed
                                            varDoNext2 = 0;
                                            if (varSessionNumber.IndexOf("*" + dbPalmClientDexId + "*") > -1)
                                                varDoNext2 = 1;

                                            varSessionNumber += "*" + dbPalmClientDexId + "*";

                                            //if (dbClient == "PLM31076")
                                            //    varTest += "DEX Id: " + dbPalmClientDexId + "\r\n";

                                            // Only do if session not processed
                                            if (varDoNext2 == 0)
                                            {
                                                // Build Activity here

                                                // Session assessments for CHSP                          
                                                if (varDoCHSP == true)
                                                {
                                                    // Set string to any value to prevent index search errors
                                                    varPreGroup = "a";
                                                    varPostGroup = "a";

                                                    // Reset session variables
                                                    sbSessionAssessment2List.Length = 0;
                                                    varSeeClient3 = false;

                                                    //ER Score loop
                                                    foreach (var s in result8.Entities)
                                                    {
                                                        // Process the data as follows:
                                                        // If there is a formatted value for the field, use it
                                                        // Otherwise if there is a literal value for the field, use it
                                                        // Otherwise the value wasn't returned so set as nothing
                                                        if (s.FormattedValues.Contains("new_palmclientdexid"))
                                                            dbPalmClientDexId2 = s.FormattedValues["new_palmclientdexid"];
                                                        else if (s.Attributes.Contains("new_palmclientdexid"))
                                                            dbPalmClientDexId2 = s.Attributes["new_palmclientdexid"].ToString();
                                                        else
                                                            dbPalmClientDexId2 = "";

                                                        if (s.FormattedValues.Contains("new_palmclientdexscores1.new_palmclientdexscoresid"))
                                                            dbPalmClientDexScoresId = s.FormattedValues["new_palmclientdexscores1.new_palmclientdexscoresid"];
                                                        else if (s.Attributes.Contains("new_palmclientdexscores1.new_palmclientdexscoresid"))
                                                            dbPalmClientDexScoresId = s.GetAttributeValue<AliasedValue>("new_palmclientdexscores1.new_palmclientdexscoresid").Value.ToString();
                                                        else
                                                            dbPalmClientDexScoresId = "";

                                                        if (s.FormattedValues.Contains("new_palmclientdexscores1.new_score"))
                                                            dbScore = s.FormattedValues["new_palmclientdexscores1.new_score"];
                                                        else if (s.Attributes.Contains("new_palmclientdexscores1.new_score"))
                                                            dbScore = s.Attributes["new_palmclientdexscores1.new_score"].ToString();
                                                        else
                                                            dbScore = "";

                                                        if (s.FormattedValues.Contains("new_palmclientdexscores1.new_section"))
                                                            dbSection = s.FormattedValues["new_palmclientdexscores1.new_section"];
                                                        else if (s.Attributes.Contains("new_palmclientdexscores1.new_section"))
                                                            dbSection = s.Attributes["new_palmclientdexscores1.new_section"].ToString();
                                                        else
                                                            dbSection = "";

                                                        // Get the section part of the field
                                                        if (String.IsNullOrEmpty(dbSection) == false)
                                                        {
                                                            dbSection = dbSection.ToLower();
                                                            if (dbSection.IndexOf(" - ") > -1)
                                                                varSection = dbSection.Substring(dbSection.IndexOf(" - ") + 3, dbSection.Length - dbSection.IndexOf(" - ") - 3);
                                                            else
                                                                varSection = "...";
                                                        }
                                                        else
                                                        {
                                                            dbSection = "...";
                                                            varSection = "...";
                                                        }

                                                        if (s.FormattedValues.Contains("new_palmclientdexscores1.new_dexrecord"))
                                                            dbDexRecord = s.FormattedValues["new_palmclientdexscores1.new_dexrecord"];
                                                        else if (s.Attributes.Contains("new_palmclientdexscores1.new_dexrecord"))
                                                            dbDexRecord = s.GetAttributeValue<AliasedValue>("new_palmclientdexscores1.new_dexrecord").Value.ToString();
                                                        else
                                                            dbDexRecord = "";

                                                        if (s.FormattedValues.Contains("new_palmclientdexscores1.new_thetype"))
                                                            dbTheType = s.FormattedValues["new_palmclientdexscores1.new_thetype"];
                                                        else if (s.Attributes.Contains("new_palmclientdexscores1.new_thetype"))
                                                            dbTheType = s.Attributes["new_palmclientdexscores1.new_thetype"].ToString();
                                                        else
                                                            dbTheType = "";

                                                        // Only do if the dex record ids match
                                                        if (dbPalmClientDexId == dbPalmClientDexId2)
                                                        {
                                                            // Score group section
                                                            if (dbSection.IndexOf("scoregroup") > -1)
                                                            {
                                                                //Get the values from the drop down tables
                                                                foreach (var d in result3.Entities)
                                                                {
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

                                                                    if (d.FormattedValues.Contains("new_er"))
                                                                        varDEX = d.FormattedValues["new_er"];
                                                                    else if (d.Attributes.Contains("new_er"))
                                                                        varDEX = d.Attributes["new_er"].ToString();
                                                                    else
                                                                        varDEX = "";

                                                                    // Append the section and score to the pre / post group section accordingly
                                                                    if (varSection.ToLower() == varDesc.ToLower())
                                                                    {
                                                                        dbSection = dbSection.Replace(varDesc.ToLower(), varDEX);
                                                                        dbSection = dbSection.Replace("scoregroup - ", "");

                                                                        if (dbTheType.ToLower() == "pre")
                                                                            varPreGroup += "," + dbSection + dbScore;
                                                                        if (dbTheType.ToLower() == "post")
                                                                            varPostGroup += "," + dbSection + dbScore;
                                                                        break;
                                                                    }

                                                                } //drop down Loop
                                                            }

                                                        } //Same client

                                                        varSeeClient3 = true; // record found

                                                    } //Scores Loop

                                                    // If the pre group was appended to above, add to extract
                                                    if (varPreGroup != "a")
                                                    {

                                                        sbSessionAssessment2List.AppendLine("         <Assessment>");

                                                        sbSessionAssessment2List.AppendLine("            <ScoreTypeCode>GROUP</ScoreTypeCode>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("            <AssessmentPhaseCode>PRE</AssessmentPhaseCode>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                        // Loop through string to get scores
                                                        value = varPreGroup.Split(',');

                                                        foreach (string dc in value)
                                                        {
                                                            if (dc != "a")
                                                                sbSessionAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                        }

                                                        sbSessionAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("         </Assessment>");

                                                    } //"a"

                                                    // If the post group was appended to above, add to extract
                                                    if (varPostGroup != "a")
                                                    {

                                                        sbSessionAssessment2List.AppendLine("         <Assessment>");

                                                        sbSessionAssessment2List.AppendLine("            <ScoreTypeCode>GROUP</ScoreTypeCode>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("            <AssessmentPhaseCode>POST</AssessmentPhaseCode>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                        // Loop through string to get scores
                                                        value = varPostGroup.Split(',');

                                                        foreach (string dc in value)
                                                        {
                                                            if (dc != "a")
                                                                sbSessionAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                        }

                                                        sbSessionAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                        sbSessionAssessment2List.AppendLine("         </Assessment>");

                                                    } //"a"

                                                    // If the assessment 2 list was appended to above, add the session assessment to the extract
                                                    if (sbSessionAssessment2List.Length > 0)
                                                    {

                                                        sbSessionAssessmentList.AppendLine("   <SessionAssessment>");
                                                        sbSessionAssessmentList.AppendLine("      <CaseId>" + varPalmClientSupportId + "_" + varExt + "</CaseId>"); //Mandatory
                                                        sbSessionAssessmentList.AppendLine("      <SessionId>" + varPalmClientDexId + "_" + varExt + "</SessionId>"); //Mandatory

                                                        sbSessionAssessmentList.AppendLine("      <Assessments>");

                                                        sbSessionAssessmentList.AppendLine(sbSessionAssessment2List.ToString());

                                                        sbSessionAssessmentList.AppendLine("      </Assessments>");
                                                        sbSessionAssessmentList.AppendLine("   </SessionAssessment>");

                                                    }

                                                } //CHSP


                                                // Client Assessments
                                                varPreCirc = "a"; // Create default data
                                                varPostCirc = "a"; // Create default data
                                                varPreGoal = "a"; // Create default data
                                                varPostGoal = "a"; // Create default data
                                                varSatisfaction = "a"; // Create default data

                                                // Reset assessment variables
                                                sbClientAssessment2List.Length = 0;
                                                varSeeClient3 = false;


                                                //ER Score loop
                                                foreach (var s in result8.Entities)
                                                {
                                                    // Process the data as follows:
                                                    // If there is a formatted value for the field, use it
                                                    // Otherwise if there is a literal value for the field, use it
                                                    // Otherwise the value wasn't returned so set as nothing
                                                    if (s.FormattedValues.Contains("new_palmclientdexid"))
                                                        dbPalmClientDexId2 = s.FormattedValues["new_palmclientdexid"];
                                                    else if (s.Attributes.Contains("new_palmclientdexid"))
                                                        dbPalmClientDexId2 = s.Attributes["new_palmclientdexid"].ToString();
                                                    else
                                                        dbPalmClientDexId2 = "";

                                                    if (s.FormattedValues.Contains("new_palmclientdexscores1.new_palmclientdexscoresid"))
                                                        dbPalmClientDexScoresId = s.FormattedValues["new_palmclientdexscores1.new_palmclientdexscoresid"];
                                                    else if (s.Attributes.Contains("new_palmclientdexscores1.new_palmclientdexscoresid"))
                                                        dbPalmClientDexScoresId = s.GetAttributeValue<AliasedValue>("new_palmclientdexscores1.new_palmclientdexscoresid").Value.ToString();
                                                    else
                                                        dbPalmClientDexScoresId = "";

                                                    if (s.FormattedValues.Contains("new_palmclientdexscores1.new_score"))
                                                        dbScore = s.FormattedValues["new_palmclientdexscores1.new_score"];
                                                    else if (s.Attributes.Contains("new_palmclientdexscores1.new_score"))
                                                        dbScore = s.Attributes["new_palmclientdexscores1.new_score"].ToString();
                                                    else
                                                        dbScore = "";

                                                    if (s.FormattedValues.Contains("new_palmclientdexscores1.new_section"))
                                                        dbSection = s.FormattedValues["new_palmclientdexscores1.new_section"];
                                                    else if (s.Attributes.Contains("new_palmclientdexscores1.new_section"))
                                                        dbSection = s.Attributes["new_palmclientdexscores1.new_section"].ToString();
                                                    else
                                                        dbSection = "";

                                                    // Get the section part of the field
                                                    if (String.IsNullOrEmpty(dbSection) == false)
                                                    {
                                                        dbSection = dbSection.ToLower();
                                                        if (dbSection.IndexOf(" - ") > -1)
                                                            varSection = dbSection.Substring(dbSection.IndexOf(" - ") + 3, dbSection.Length - dbSection.IndexOf(" - ") - 3);
                                                        else
                                                            varSection = "...";
                                                    }
                                                    else
                                                    {
                                                        dbSection = "...";
                                                        varSection = "...";
                                                    }

                                                    if (s.FormattedValues.Contains("new_palmclientdexscores1.new_dexrecord"))
                                                        dbDexRecord = s.FormattedValues["new_palmclientdexscores1.new_dexrecord"];
                                                    else if (s.Attributes.Contains("new_palmclientdexscores1.new_dexrecord"))
                                                        dbDexRecord = s.GetAttributeValue<AliasedValue>("new_palmclientdexscores1.new_dexrecord").Value.ToString();
                                                    else
                                                        dbDexRecord = "";

                                                    if (s.FormattedValues.Contains("new_palmclientdexscores1.new_thetype"))
                                                        dbTheType = s.FormattedValues["new_palmclientdexscores1.new_thetype"];
                                                    else if (s.Attributes.Contains("new_palmclientdexscores1.new_thetype"))
                                                        dbTheType = s.Attributes["new_palmclientdexscores1.new_thetype"].ToString();
                                                    else
                                                        dbTheType = "";

                                                    // We need to get the entity id for the support period field for comparisons
                                                    if (s.Attributes.Contains("new_financial"))
                                                    {
                                                        // Get the entity id for the client using the entity reference object
                                                        getEntity = (EntityReference)s.Attributes["new_financial"];
                                                        dbSupportPeriod2 = getEntity.Id.ToString();
                                                    }
                                                    else if (s.FormattedValues.Contains("new_financial"))
                                                        dbSupportPeriod2 = s.FormattedValues["new_financial"];
                                                    else
                                                        dbSupportPeriod2 = "";

                                                    //if (dbClient == "PLM31076")
                                                    //    varTest += dbSupportPeriod2 + " " + dbPalmClientDexId2 + " " + dbDexRecord + " " + dbTheType + " " + dbSection + " " + dbScore + "\r\n";

                                                    // Only process if the DEX ids match
                                                    if (dbPalmClientDexId == dbPalmClientDexId2)
                                                    {
                                                        if (dbSection.IndexOf("scorecircumstances") > -1 || dbSection.IndexOf("scoregoals") > -1 || dbSection.IndexOf("scoresatisfaction") > -1)
                                                        {
                                                            //Get the values from the drop down tables
                                                            foreach (var d in result3.Entities)
                                                            {
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

                                                                if (d.FormattedValues.Contains("new_er"))
                                                                    varDEX = d.FormattedValues["new_er"];
                                                                else if (d.Attributes.Contains("new_er"))
                                                                    varDEX = d.Attributes["new_er"].ToString();
                                                                else
                                                                    varDEX = "";

                                                                // Process if the section matches and the type is valid
                                                                if (varSection.ToLower() == varDesc.ToLower() && (varType == "scorecircumstances" || varType == "scoregoals" || varType == "scoresatisfaction"))
                                                                {
                                                                    //if (dbClient == "PLM31076")
                                                                    //    varTest += "*" + varSection + "* *" + varType + " " + varDesc + "* *" + varDEX + "*\r\n";

                                                                    dbSection = dbSection.Replace(varDesc.ToLower(), varDEX);

                                                                    // Append the section and score to the pre / post goals section accordingly
                                                                    if (dbSection.IndexOf("scoregoals") > -1)
                                                                    {
                                                                        dbSection = dbSection.Replace("scoregoals - ", "");
                                                                        if (dbTheType.ToLower() == "pre")
                                                                        {
                                                                            varPreGoal += "," + dbSection + dbScore;
                                                                        }
                                                                        if (dbTheType.ToLower() == "post")
                                                                        {
                                                                            varPostGoal += "," + dbSection + dbScore;
                                                                        }
                                                                    }
                                                                    // Append the section and score to the pre / post circumstances section accordingly
                                                                    if (dbSection.IndexOf("scorecircumstances") > -1)
                                                                    {
                                                                        dbSection = dbSection.Replace("scorecircumstances - ", "");
                                                                        if (dbTheType.ToLower() == "pre")
                                                                        {
                                                                            varPreCirc += "," + dbSection + dbScore;
                                                                        }
                                                                        if (dbTheType.ToLower() == "post")
                                                                        {
                                                                            varPostCirc += "," + dbSection + dbScore;
                                                                        }
                                                                    }
                                                                    // Append the section and score to the post satisfaction section accordingly
                                                                    if (dbSection.IndexOf("scoresatisfaction") > -1)
                                                                    {
                                                                        dbSection = dbSection.Replace("scoresatisfaction - ", "");
                                                                        if (dbTheType.ToLower() == "post")
                                                                        {
                                                                            varSatisfaction += "," + dbSection + dbScore;
                                                                        }
                                                                    }
                                                                    break;
                                                                }

                                                            } //drop down Loop
                                                        }

                                                    } //Same client

                                                    varSeeClient3 = true; // Client record found

                                                } //Scores Loop


                                                // If the pre circ was appended to above, add to extract
                                                if (varPreCirc != "a")
                                                {

                                                    sbClientAssessment2List.AppendLine("         <Assessment>");

                                                    sbClientAssessment2List.AppendLine("            <ScoreTypeCode>CIRCUMSTANCES</ScoreTypeCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <AssessmentPhaseCode>PRE</AssessmentPhaseCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                    // Loop through string to get scores
                                                    value = varPreCirc.Split(',');

                                                    foreach (string dc in value)
                                                    {
                                                        if (dc != "a")
                                                            sbClientAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                    }

                                                    sbClientAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("         </Assessment>");

                                                } //"a"

                                                // If the post circ was appended to above, add to extract
                                                if (varPostCirc != "a")
                                                {

                                                    sbClientAssessment2List.AppendLine("         <Assessment>");

                                                    sbClientAssessment2List.AppendLine("            <ScoreTypeCode>CIRCUMSTANCES</ScoreTypeCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <AssessmentPhaseCode>POST</AssessmentPhaseCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                    // Loop through string to get scores
                                                    value = varPostCirc.Split(',');

                                                    foreach (string dc in value)
                                                    {
                                                        if (dc != "a")
                                                            sbClientAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                    }

                                                    sbClientAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("         </Assessment>");

                                                } //"a"

                                                // If the pre goal was appended to above, add to extract
                                                if (varPreGoal != "a")
                                                {

                                                    sbClientAssessment2List.AppendLine("         <Assessment>");

                                                    sbClientAssessment2List.AppendLine("            <ScoreTypeCode>GOALS</ScoreTypeCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <AssessmentPhaseCode>PRE</AssessmentPhaseCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                    // Loop through string to get scores
                                                    value = varPreGoal.Split(',');

                                                    foreach (string dc in value)
                                                    {
                                                        if (dc != "a")
                                                            sbClientAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                    }

                                                    sbClientAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("         </Assessment>");

                                                } //"a"

                                                // If the post goal was appended to above, add to extract
                                                if (varPostGoal != "a")
                                                {

                                                    sbClientAssessment2List.AppendLine("         <Assessment>");

                                                    sbClientAssessment2List.AppendLine("            <ScoreTypeCode>GOALS</ScoreTypeCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <AssessmentPhaseCode>POST</AssessmentPhaseCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                    // Loop through string to get scores
                                                    value = varPostGoal.Split(',');

                                                    foreach (string dc in value)
                                                    {
                                                        if (dc != "a")
                                                            sbClientAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                    }

                                                    sbClientAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("         </Assessment>");

                                                } //"a"

                                                // If the satisfaction was appended to above, add to extract
                                                if (varSatisfaction != "a")
                                                {

                                                    sbClientAssessment2List.AppendLine("         <Assessment>");

                                                    sbClientAssessment2List.AppendLine("            <ScoreTypeCode>SATISFACTION</ScoreTypeCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <AssessmentPhaseCode>POST</AssessmentPhaseCode>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("            <Scores>"); //Mandatory

                                                    // Loop through string to get scores
                                                    value = varSatisfaction.Split(',');

                                                    foreach (string dc in value)
                                                    {
                                                        if (dc != "a")
                                                            sbClientAssessment2List.AppendLine("               <ScoreCode>" + dc + "</ScoreCode>"); //Mandatory
                                                    }

                                                    sbClientAssessment2List.AppendLine("            </Scores>"); //Mandatory
                                                    sbClientAssessment2List.AppendLine("         </Assessment>");

                                                } //"a"

                                                // If the assessment 2 list was appended to above, add the client assessment to the extract
                                                if (sbClientAssessment2List.Length > 0)
                                                {

                                                    sbClientAssessmentList.AppendLine("   <ClientAssessment>");
                                                    sbClientAssessmentList.AppendLine("      <ClientId>" + dbExternal + "</ClientId>"); //Mandatory
                                                    sbClientAssessmentList.AppendLine("      <CaseId>" + varPalmClientSupportId + "_" + varExt + "</CaseId>"); //Mandatory
                                                    sbClientAssessmentList.AppendLine("      <SessionId>" + varPalmClientDexId + "_" + varExt + "</SessionId>"); //Mandatory

                                                    sbClientAssessmentList.AppendLine("      <Assessments>");

                                                    sbClientAssessmentList.AppendLine(sbClientAssessment2List.ToString());

                                                    sbClientAssessmentList.AppendLine("      </Assessments>");
                                                    sbClientAssessmentList.AppendLine("   </ClientAssessment>");

                                                }

                                                // Get the service type id
                                                varServTypeId = "1";
                                                if (varDoCHSP == true)
                                                    varServTypeId = "192";

                                                //Get the assistance type and convert to servicetype id - result3 = dropdownlist.
                                                foreach (var d in result3.Entities)
                                                {
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

                                                    if (d.FormattedValues.Contains("new_er"))
                                                        varDEX = d.FormattedValues["new_er"];
                                                    else if (d.Attributes.Contains("new_er"))
                                                        varDEX = d.Attributes["new_er"].ToString();
                                                    else
                                                        varDEX = "";

                                                    if (varDesc.ToLower() == dbServType.ToLower() && varType == "erassist")
                                                    {
                                                        varServTypeId = varDEX; //20 is - accom, bonds, ooh bond debt, rent arrears, rent in advance.
                                                                                //varDebug = varDesc;
                                                        break;
                                                    }

                                                } //k Loop

                                                // Build Session here                              
                                                sbSessionList.AppendLine("   <Session>");

                                                //Mandatory String
                                                sbSessionList.AppendLine("      <SessionId>" + varPalmClientDexId + "_" + varExt + "</SessionId>");

                                                //Mandatory String
                                                sbSessionList.AppendLine("      <CaseId>" + varPalmClientSupportId + "_" + varExt + "</CaseId>");

                                                //Mandatory Date
                                                sbSessionList.AppendLine("      <SessionDate>" + dbEntryDate + "</SessionDate>");

                                                //ToDo: Optional Integer
                                                sbSessionList.AppendLine("      <ServiceTypeId>" + varServTypeId + "</ServiceTypeId>");

                                                //Optional Integer
                                                if (varGetRecip > 0 && 1 == 2)
                                                    sbSessionList.AppendLine("      <TotalNumberOfUnidentifiedClients>" + varGetRecip + "</TotalNumberOfUnidentifiedClients>"); //Mandatory
                                                else
                                                    sbSessionList.AppendLine("      <TotalNumberOfUnidentifiedClients>0</TotalNumberOfUnidentifiedClients>"); //Mandatory

                                                // For CHSP only
                                                //if (varDoCHSP == true)
                                                //{
                                                //  sbSessionList.AppendLine("      <FeesCharged>0.00</FeesCharged>");
                                                //}

                                                //Optional True / False
                                                //sbSessionList.AppendLine("      <InterpreterPresent>" + varInterpreterUsed + "</InterpreterPresent>");


                                                sbSessionList.AppendLine("      <SessionClients>");
                                                sbSessionList.AppendLine("         <SessionClient>");
                                                sbSessionList.AppendLine("            <ClientId>" + dbExternal + "</ClientId>");
                                                //sbSessionList.AppendLine ("            <ParticipationTypeCode>CLIENT</ParticipationTypeCode>");
                                                sbSessionList.AppendLine("            <ParticipationCode>CLIENT</ParticipationCode>");
                                                sbSessionList.AppendLine("         </SessionClient>");
                                                sbSessionList.AppendLine("      </SessionClients>");

                                                // For CHSP only
                                                if (varDoCHSP == true)
                                                    sbSessionList.AppendLine("      <TimeMinutes>" + dbHours + "</TimeMinutes>");

                                                sbSessionList.AppendLine("   </Session>");

                                            } // varDoNext2

                                        } // Same support period

                                    } // Dex Loop

                                } // prevent dup case

                            } // client loop


                            //Header part of the DEX extract
                            sbHeaderList.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");

                            sbHeaderList.AppendLine("<DEXFileUpload>");

                            //Section 1: Clients
                            sbHeaderList.AppendLine("<Clients>");
                            sbHeaderList.AppendLine(sbClientList.ToString());
                            sbHeaderList.AppendLine("</Clients>");

                            //Section 2: Cases
                            sbHeaderList.AppendLine("<Cases>");
                            sbHeaderList.AppendLine(sbCaseList.ToString());
                            sbHeaderList.AppendLine("</Cases>");

                            //Section 3: Sessions
                            sbHeaderList.AppendLine("<Sessions>");
                            sbHeaderList.AppendLine(sbSessionList.ToString());
                            sbHeaderList.AppendLine("</Sessions>");

                            //Section 4: Session Assessments | Group Only
                            if (sbSessionAssessmentList.Length > 0)
                            {
                                sbHeaderList.AppendLine("<SessionAssessments>");
                                sbHeaderList.AppendLine(sbSessionAssessmentList.ToString());
                                sbHeaderList.AppendLine("</SessionAssessments>");
                            }

                            //Section 5: Client Assessments
                            sbHeaderList.AppendLine("<ClientAssessments>");
                            sbHeaderList.AppendLine(sbClientAssessmentList.ToString());
                            sbHeaderList.AppendLine("</ClientAssessments>");

                            sbHeaderList.AppendLine("</DEXFileUpload>");


                            //varTest += sbHeaderList.ToString();

                            // Create note against current Palm Go DEX record and add attachment
                            byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                            string encodedData = System.Convert.ToBase64String(filename);
                            Entity Annotation = new Entity("annotation");
                            Annotation.Attributes["objectid"] = new EntityReference("new_palmgodex", varDexID);
                            Annotation.Attributes["objecttypecode"] = "new_palmgodex";
                            Annotation.Attributes["subject"] = "DEX Extract";
                            Annotation.Attributes["documentbody"] = encodedData;
                            Annotation.Attributes["mimetype"] = @"text / plain";
                            Annotation.Attributes["notetext"] = "DEX Extract for " + cleanDate(varStartDatePr) + " to " + cleanDate(varEndDatePr);
                            Annotation.Attributes["filename"] = varFileName;
                            _service.Create(Annotation);

                            // If there is an error, create note against current Palm Go DEX record and add attachment
                            if (sbErrorList.Length > 0)
                            {
                                byte[] filename2 = Encoding.ASCII.GetBytes(sbErrorList.ToString());
                                string encodedData2 = System.Convert.ToBase64String(filename2);
                                Entity Annotation2 = new Entity("annotation");
                                Annotation2.Attributes["objectid"] = new EntityReference("new_palmgodex", varDexID);
                                Annotation2.Attributes["objecttypecode"] = "new_palmgodex";
                                Annotation2.Attributes["subject"] = "DEX Extract";
                                Annotation2.Attributes["documentbody"] = encodedData2;
                                Annotation2.Attributes["mimetype"] = @"text / plain";
                                Annotation2.Attributes["notetext"] = "DEX errors and warnings for " + cleanDate(varStartDatePr) + " to " + cleanDate(varEndDatePr);
                                Annotation2.Attributes["filename"] = varFileName2;
                                _service.Create(Annotation2);
                            }

                            //varTest += cleanDate(varStartDatePr) + " " + cleanDate(varEndDatePr);

                            //throw new InvalidPluginExecutionException("This plugin is working\r\n" + varTest);
                        }
                    }


                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in the FollowupPlugin plug-in. :  ex:" + ex.ToString(), ex);
                }
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

        //Format: 01/01/1970
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

        //Format: 1-Jan-1970
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

        //Round up number
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

        //Format: 1970-01-01
        public string cleanDateE(DateTime getdate)
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

            string clean = getdate.Year + "-" + varMonth + "-" + varDay;
            return clean;
        }

        // Date format for DEX with year only
        public string cleanDateEs(DateTime getdate)
        {
            string clean = getdate.Year + "-01-01";
            return clean;
        }

        //Date format for SHOR: 01062013
        public string cleanDateS(DateTime getdate)
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


            string clean = varDay + varMonth + getdate.Year;
            return clean;
        }
    }
}

