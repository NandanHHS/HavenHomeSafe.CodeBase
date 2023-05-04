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
    public class goSHOR : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins
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
                // if (entity.LogicalName != "new_palmgoshor")
                //    return;

                try
                {
                    // Fixes:
                    // Remove ampersand from comments
                    // Compare values to Lower
                    // Use drop down list values instead of hardcode
                    // Error if there is a value but not in matching dropdown list

                    string varDescription = ""; // Description field on form
                    bool varCreateExtract = false; // Create Extract field on form
                    string varAgencyId = ""; // Agency Id on form
                    string varAgencyName = ""; // Agency Name on form
                    Guid varShorID = new Guid(); // GUID for palm go shor record
                    StringBuilder sbExtractList = new StringBuilder(); // String builder for extract
                    StringBuilder sbErrorList = new StringBuilder(); // String builder for errors
                    StringBuilder sbSHORHomeSt = new StringBuilder(); // String builder for homeless status update
                    StringBuilder sbSHORHomeMonth = new StringBuilder(); // String builder for homeless month
                    StringBuilder sbSHORHomeYear = new StringBuilder(); // String builder for homeless year
                    StringBuilder sbSHORFac = new StringBuilder(); // String builder for facilities
                    StringBuilder sbSHORReas = new StringBuilder(); // String builder for reasons
                    StringBuilder sbSHORAccomServ = new StringBuilder(); // String builder for accom
                    StringBuilder sbSHORTAR = new StringBuilder(); // String builder for turnaway reasons
                    StringBuilder sbSHORTAQ = new StringBuilder(); // String builder for turnaway request
                    StringBuilder sbHeaderList = new StringBuilder(); // String builder for extract header
                    StringBuilder sbSHORServices = new StringBuilder(); // String builder for services
                    StringBuilder sbClientList = new StringBuilder(); // String builder for clients
                    DateTime varMonthStart = new DateTime(); // Period Start
                    DateTime varMonthEnd = new DateTime(); // Period End
                    string varPrintMonth = ""; // String for period
                    DateTime varPrevMonthSt = new DateTime(); // Previous period start
                    DateTime varPrevMonthEn = new DateTime(); // Previous period end
                    DateTime varCheckDate = new DateTime(); // Used to parse dates
                    DateTime varCurrentDate = DateTime.Now; // Current date
                    double varCheckDouble = 0; // Used to parse doubles
                    int varCheckInt = 0; // Used to parse integers
                    bool varDoErr = false; // Whether an error has occured
                    bool varSeeType = false; // Whether financial types was found

                    string varDebug = "";

                    EntityReference getEntity; // Entity reference object

                    varMonthStart = DateTime.Now; // Set start month to current date

                    string varTest = ""; // Debug

                    // Only do this if the entity is the Palm Go SHOR entity
                    if (entity.LogicalName == "new_palmgoshor")
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

                        try
                        {
                            // Get info from current Palm Go SHOR record
                            varDescription = entity.GetAttributeValue<string>("new_description");
                            varCreateExtract = entity.GetAttributeValue<bool>("new_createextract");
                            varMonthStart = entity.GetAttributeValue<DateTime>("new_startdate");
                            varShorID = entity.Id;

                            // Get the SHOR agency information from an entity reference object
                            EntityReference ownerLookup = (EntityReference)entity.Attributes["new_agencyid"];
                            varAgencyId += ownerLookup.Id.ToString() + ".\r\n";
                            varAgencyId += ((EntityReference)entity.Attributes["new_agencyid"]).Name + ".\r\n";
                            varAgencyId += ownerLookup.LogicalName + ".\r\n";

                            var actualOwningUnit = _service.Retrieve(ownerLookup.LogicalName, ownerLookup.Id, new ColumnSet(true));

                            // Assign agency name and id
                            varAgencyId = actualOwningUnit["new_agencyid"].ToString();
                            varAgencyName = actualOwningUnit["new_agency"].ToString();

                            // Important: The plugin uses American dates but returns formatted Australian dates
                            // Any dates created in the plugin will be American
                            // Add 14 hours to current month
                            varMonthStart = varMonthStart.AddHours(14);

                            // Get month start and month end
                            varMonthStart = Convert.ToDateTime("1-" + varMonthStart.ToString("MMM") + "-" + varMonthStart.Year);
                            varMonthEnd = varMonthStart.AddMonths(1);
                            varMonthEnd = varMonthEnd.AddDays(-1).AddHours(23);
                            varTest = "1: " + varMonthStart + ", 2: " + varMonthEnd;
                            // Get month number with leading zero
                            if (varMonthStart.Month < 10)
                                varPrintMonth = "0" + varMonthStart.Month;
                            else
                                varPrintMonth = varMonthStart.Month + "";

                            //Get the previous period
                            varPrevMonthSt = varMonthStart.AddMonths(-1);
                            varPrevMonthEn = varMonthEnd.AddMonths(-1);
                            if(varPrevMonthEn.Month == 1 || varPrevMonthEn.Month == 3|| varPrevMonthEn.Month == 5|| varPrevMonthEn.Month == 7 || varPrevMonthEn.Month == 8|| varPrevMonthEn.Month == 10|| varPrevMonthEn.Month == 12)
                            {
                                varPrevMonthEn = new DateTime(varPrevMonthEn.Year, varPrevMonthEn.Month, 31);
                            }
                            else if(varPrevMonthEn.Month == 2)
                            {
                                if(varPrevMonthEn.Year % 4 == 0 && (varPrevMonthEn.Year % 100 != 0 || varPrevMonthEn.Year % 400 == 0))
                                    varPrevMonthEn = new DateTime(varPrevMonthEn.Year, varPrevMonthEn.Month, 29);
                                else
                                    varPrevMonthEn = new DateTime(varPrevMonthEn.Year, varPrevMonthEn.Month, 28);
                            }
                            else
                            {
                                varPrevMonthEn = new DateTime(varPrevMonthEn.Year, varPrevMonthEn.Month, 30);
                            }

                            //C02.001.02 | Error if date less than July 2011
                            if (varMonthStart < Convert.ToDateTime("1-Jul-2012"))
                                sbErrorList.AppendLine("Error: Extract prior to July 2012 cannot be produced");

                            //C02.001.03 | Error if date greater than current date
                            if (varMonthStart > DateTime.Now)
                                sbErrorList.AppendLine("Error: Extract cannot be produced for a future month");

                            //C03.001.01-7 | Error if not valid agency ID or agency name (also C03.002.01-2)
                            if (String.IsNullOrEmpty(varAgencyId) == true)
                                sbErrorList.AppendLine("Error: Invalid agencyid");

                            //varTest += varMonthStart + " " + varMonthEnd + "\r\n";

                            // Fetch statements for database
                            // Get the required fields from the client table (and associated entities)
                            // Any clients that have a support period ticked as SHS record for the current period
                            string dbSHORList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclient'>
                                <attribute name='new_address' />
                                <attribute name='new_firstname' />
                                <attribute name='new_surname' />
                                <attribute name='new_gender' />
                                <attribute name='new_sex' />
                                <attribute name='new_dob' />
                                <attribute name='new_dobest' />
                                <attribute name='new_dobestd' />
                                <attribute name='new_dobestm' />
                                <attribute name='new_dobesty' />
                                <attribute name='new_yeararrival' />
                                <attribute name='new_indigenous' />
                                <attribute name='new_country' />
                                <attribute name='new_consent' />
                                <attribute name='new_language' />
                                <attribute name='new_languageother' />
                                <attribute name='new_speakproficient' />
                                <attribute name='createdon' />
                                <link-entity name='new_palmclientsupport' to='new_palmclientid' from='new_client' link-type='inner'>
                                    <attribute name='new_adf' />
                                    <attribute name='new_assessed' />
                                    <attribute name='new_awaitgovtpres' />
                                    <attribute name='new_awaitgovtwkbef' />
                                    <attribute name='new_carepres' />
                                    <attribute name='new_carewkbef' />
                                    <attribute name='new_cessation' />
                                    <attribute name='new_diagnosedmh' />
                                    <attribute name='new_discomm' />
                                    <attribute name='new_dismob' />
                                    <attribute name='new_disself' />
                                    <attribute name='new_educationpres' />
                                    <attribute name='new_facilities' />
                                    <attribute name='new_ftptpres' />
                                    <attribute name='new_ftptwkbef' />
                                    <attribute name='new_homemonth' />
                                    <attribute name='new_homeyear' />
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_supportperiodidold' />
                                    <attribute name='new_puhidold' />
                                    <attribute name='new_incomepres' />
                                    <attribute name='new_incomewkbef' />
                                    <attribute name='new_labourforcepres' />
                                    <attribute name='new_labourforcewkbef' />
                                    <attribute name='new_livingarrangepres' />
                                    <attribute name='new_livingarrangewkbef' />
                                    <attribute name='new_mentalillnessinfo' />
                                    <attribute name='new_mhservicesrecd' />
                                    <attribute name='new_newclient' />
                                    <attribute name='new_occupancypres' />
                                    <attribute name='new_occupancywkbef' />
                                    <attribute name='new_preslocality' />
                                    <attribute name='new_primreason' />
                                    <attribute name='new_program' />
                                    <attribute name='new_puhid' />
                                    <attribute name='new_puhrship' />
                                    <attribute name='new_puhrshipother' />
                                    <attribute name='new_reasons' />
                                    <attribute name='new_reasonsoth' />
                                    <attribute name='new_residentialpres' />
                                    <attribute name='new_residentialwkbef' />
                                    <attribute name='new_shordate' />
                                    <attribute name='new_sourceref' />
                                    <attribute name='new_startdate' />
                                    <attribute name='new_enddate' />
                                    <attribute name='new_studindpres' />
                                    <attribute name='new_studindwkbef' />
                                    <attribute name='new_studtypepres' />
                                    <attribute name='new_studtypewkbef' />
                                    <attribute name='new_tenurepres' />
                                    <attribute name='new_tenurewkbef' />
                                    <attribute name='new_timeperm' />
                                    <attribute name='ownerid' />
                                    <attribute name='new_wkbeflocality' />
                                    <attribute name='new_description' />
                                    <attribute name='new_ndis' />
                                    <link-entity name='new_palmddllocality' to='new_preslocality' from='new_palmddllocalityid' link-type='outer'>
                                        <attribute name='new_postcode' />
                                        <attribute name='new_state' />
                                    </link-entity>
                                    <link-entity name='new_palmddllocality' to='new_wkbeflocality' from='new_palmddllocalityid' link-type='outer'>
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
                                <link-entity name='new_palmddllanguage' to='new_languageother' from='new_palmddllanguageid' link-type='outer'>
                                    <attribute name='new_code' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientsupport' attribute='new_shor' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='lt' value='" + varMonthEnd + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_doshor' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + varMonthStart + @"' />
                                    </filter >
                                </filter>
                                <order attribute='new_address' />
                              </entity>
                            </fetch> ";

                            // Get the required fields from the status update table
                            // Any status updates for the current period
                            string dbStatusList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientshorstatus'>
                                <attribute name='new_awaitgovt' />
                                <attribute name='new_careorder' />
                                <attribute name='new_casemgt' />
                                <attribute name='new_casemgtgoal' />
                                <attribute name='new_casemgtreason' />
                                <attribute name='new_casemgtreasonoth' />
                                <attribute name='new_entrydate' />
                                <attribute name='new_ftpt' />
                                <attribute name='new_homeless' />
                                <attribute name='new_income' />
                                <attribute name='new_labourforce' />
                                <attribute name='new_livingarrange' />
                                <attribute name='new_occupancy' />
                                <attribute name='new_residential' />
                                <attribute name='new_studind' />
                                <attribute name='new_studtype' />
                                <attribute name='new_supportperiod' />
                                <attribute name='new_tenure' />
                                <attribute name='new_resubmit' />
                                <attribute name='new_endreason' />
                                <attribute name='new_ongoing' />
                                <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                    <attribute name='new_doshor' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientshorstatus' attribute='new_entrydate' operator='le' value='" + varMonthEnd + @"' />
                                    <condition entityname='new_palmclientshorstatus' attribute='new_entrydate' operator='ge' value='" + varMonthStart + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_doshor' operator='eq' value='" + ownerLookup.Id + @"' />
                                </filter >
                                <order attribute='new_supportperiod' />
                              </entity>
                            </fetch> ";

                            // Get the required fields from the status update table for previous period
                            // Any status updates for the previous period
                            string dbPrevList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientshorstatus'>
                                <attribute name='new_supportperiod' />
                                <attribute name='new_entrydate' />
                                <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                    <attribute name='new_doshor' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientshorstatus' attribute='new_entrydate' operator='le' value='" + cleanDate(varPrevMonthEn) + @"' />
                                    <condition entityname='new_palmclientshorstatus' attribute='new_entrydate' operator='ge' value='" + cleanDate(varPrevMonthSt) + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_doshor' operator='eq' value='" + ownerLookup.Id + @"' />
                                </filter >
                                <order attribute='new_supportperiod' />
                              </entity>
                            </fetch> ";

                            // Get the required fields from the SU Drop Down list entity
                            string dbDropList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                  <entity name='new_palmsudropdown'>
                                    <attribute name='new_type' />
                                    <attribute name='new_description' />
                                    <attribute name='new_shor' />
                                    <order attribute='new_description' />
                                  </entity>
                                </fetch> ";

                            // Get the required fields from the services table
                            // Any services for the current period against a support period ticked as SHS
                            string dbServicesList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                    <entity name='new_palmclientservices'>
                                        <attribute name='new_description' />
                                        <attribute name='new_entrydate' />
                                        <attribute name='new_servarrange' />
                                        <attribute name='new_servneed' />
                                        <attribute name='new_servprovide' />
                                        <attribute name='new_supportperiod' />
                                        <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                            <attribute name='new_doshor' />
                                        </link-entity>
                                        <filter type='and'>
                                            <condition entityname='new_palmclientservices' attribute='new_entrydate' operator='le' value='" + varMonthEnd + @"' />
                                            <condition entityname='new_palmclientservices' attribute='new_entrydate' operator='ge' value='" + varMonthStart + @"' />
                                            <condition entityname='new_palmclientsupport' attribute='new_doshor' operator='eq' value='" + ownerLookup.Id + @"' />
                                        </filter >
                                    </entity>
                                </fetch> ";

                            // Get the required fields from the accommodation table
                            // Any accom for the current period against a support period ticked as SHS
                            string dbAccomList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                    <entity name='new_palmclientaccom'>
                                        <attribute name='new_accomtype' />
                                        <attribute name='new_datefrom' />
                                        <attribute name='new_dateto' />
                                        <attribute name='new_location' />
                                        <attribute name='new_supportperiod' />
                                        <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                            <attribute name='new_doshor' />
                                        </link-entity>
                                        <filter type='and'>
                                            <condition entityname='new_palmclientaccom' attribute='new_datefrom' operator='le' value='" + varMonthEnd + @"' />
                                            <condition entityname='new_palmclientsupport' attribute='new_doshor' operator='eq' value='" + ownerLookup.Id + @"' />
                                            <filter type='or'>
                                                <condition entityname='new_palmclientaccom' attribute='new_dateto' operator='ge' value='" + varMonthStart + @"' />
                                                <condition entityname='new_palmclientaccom' attribute='new_dateto' operator='null' />
                                            </filter >
                                        </filter >
                                    </entity>
                                </fetch> ";

                            // Get the required fields from the financial table
                            // Any financials for the current period against a support period ticked as SHS
                            string dbFinancialList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                    <entity name='new_palmclientfinancial'>
                                        <attribute name='new_amount' />
                                        <attribute name='new_entrydate' />
                                        <attribute name='new_shor' />
                                        <attribute name='new_supportperiod' />
                                        <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                            <attribute name='new_doshor' />
                                        </link-entity>
                                        <filter type='and'>
                                            <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='le' value='" + varMonthEnd + @"' />
                                            <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + varMonthStart + @"' />
                                            <condition entityname='new_palmclientsupport' attribute='new_doshor' operator='eq' value='" + ownerLookup.Id + @"' />
                                        </filter>
                                    </entity>
                                </fetch> ";


                            // Get the required fields from the unassists table
                            // Any unassists for the current period
                            string dbStepList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                    <entity name='new_palmstep'>
                                        <attribute name='new_agencyid' />
                                        <attribute name='new_date' />
                                        <attribute name='new_dob' />
                                        <attribute name='new_dobestd' />
                                        <attribute name='new_dobestm' />
                                        <attribute name='new_dobesty' />
                                        <attribute name='new_estimated' />
                                        <attribute name='new_firstname' />
                                        <attribute name='new_firstserv' />
                                        <attribute name='new_sex' />
                                        <attribute name='new_ispalm' />
                                        <attribute name='new_isstep' />
                                        <attribute name='new_name' />
                                        <attribute name='new_relationship' />
                                        <attribute name='new_relother' />
                                        <attribute name='new_surname' />
                                        <attribute name='new_turnassist' />
                                        <attribute name='new_turnreason' />
                                        <attribute name='new_turnreasoth' />
                                        <attribute name='new_turnrequest' />
                                        <attribute name='new_turnurg' />
                                        <attribute name='new_palmstepid' />
                                        <filter type='and'>
                                            <condition entityname='new_palmstep' attribute='new_agencyid' operator='eq' value='" + ownerLookup.Id + @"' />
                                            <condition entityname='new_palmstep' attribute='new_date' operator='le' value='" + varMonthEnd + @"' />
                                            <condition entityname='new_palmstep' attribute='new_date' operator='ge' value='" + varMonthStart + @"' />
                                        </filter >
                                    </entity>
                                </fetch> ";

                            // Get a distinct list of presenting units for unassists
                            string dbTurnPuhList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                    <entity name='new_palmstep'>
                                        <attribute name='new_palmstepid' />
                                        <filter type='and'>
                                            <condition entityname='new_palmstep' attribute='new_agencyid' operator='eq' value='" + ownerLookup.Id + @"' />
                                        </filter >
                                    </entity>
                                </fetch> ";

                            // Database variables
                            string dbClient = ""; // Client field
                            string dbFirstName = ""; // First name field
                            string dbSurname = ""; // Surname field
                            string dbGender = ""; // Gender field
                            string dbSex = ""; // Sex field
                            string dbDob = ""; // DOB field
                            string dbDobEst = ""; // DOB estimated field (old)
                            string dbDobEstD = ""; // DOB day estimated field
                            string dbDobEstM = ""; // DOB month estimated field
                            string dbDobEstY = ""; // DOB year estimated field
                            string dbYearArrival = ""; // Year arrival field
                            string dbIndigenous = ""; // Indigenous field
                            string dbCountry = ""; // Country field
                            string dbConsent = ""; // Consent field
                            string dbADF = ""; // ADF field
                            string dbAssessed = ""; // Assessed field
                            string dbAwaitGovtPres = ""; // Await govt pres field
                            string dbAwaitGovtWkBef = ""; // Await govt week before field
                            string dbCarePres = ""; // Child care pres field
                            string dbCareWkBef = ""; // Child care week before field
                            string dbCessation = ""; // Cessation field
                            string dbDiagnosedMH = ""; // Diagnosed mental health field
                            string dbDisComm = ""; // Disability Communication field
                            string dbDisMob = ""; // Disability Mobility field
                            string dbDisSelf = ""; // Disability Self Care field
                            string dbEducationPres = ""; // Education field
                            string dbFacilities = ""; // Facilities field
                            string dbFtPtPres = ""; // FT PT pres field
                            string dbFtPtWkBef = ""; // FT PT week before field
                            string dbHomeMonth = ""; // Homeless month field
                            string dbHomeYear = ""; // Homeless year field
                            string dbPalmClientSupportId = ""; // Support period id field
                            string dbIncomePres = ""; // Income presenting field
                            string dbIncomeWkBef = ""; // Income week before field
                            string dbLabourForcePres = ""; // Labour force pres field
                            string dbLabourForceWkBef = ""; // Labour force week before field
                            string dbLivingArrangePres = ""; // Living arrange pres field
                            string dbLivingArrangeWkBef = ""; // Living arrange week before field
                            string dbMentalIllnessInfo = ""; // Mental illness source field
                            string dbMHServicesRecd = ""; // MH services recieved field
                            string dbNewClient = ""; // New client field
                            string dbOccupancyPres = ""; // Occupancy pres field
                            string dbOccupancyWkBef = ""; // Occupancy week before field
                            string dbPresLocality = ""; // Presenting locality field
                            string dbPresState = ""; // Presenting state field
                            string dbPresPostcode = ""; // Presenting postcode field
                            string dbPrimReason = ""; // Primary reason field
                            string dbProgram = ""; // Program field
                            string dbPuhId = ""; // Presenting unit field
                            string dbPuhRship = ""; // Presenting unit relationship field
                            string dbPuhRshipOther = ""; // Presenting unit relationship (other) field
                            string dbReasons = ""; // Reasons field
                            string dbReasonsOth = ""; // Reasons (other) field
                            string dbResidentialPres = ""; // Residential pres field
                            string dbResidentialWkBef = ""; // Residential week before field
                            string dbShorDate = ""; // Force SHS date field
                            string dbSourceRef = ""; // Source of referral field
                            string dbStartDate = ""; // Start date field
                            string dbEndDate = ""; // End date field
                            string dbStudIndPres = ""; // Student indicator pres field
                            string dbStudIndWkBef = ""; // Student indicator week before field
                            string dbStudTypePres = ""; // Student type pres field
                            string dbStudTypeWkBef = ""; // Student type week before field
                            string dbTenurePres = ""; // Tenure pres field
                            string dbTenureWkBef = ""; // Tenure week before field
                            string dbTimePerm = ""; // Time since last perm field
                            string dbOwnerId = ""; // Owner field
                            string dbWkBefLocality = ""; // Week before locality field
                            string dbWkBefState = ""; // Week before state field
                            string dbWkBefPostcode = ""; // Week before postcode field
                            string dbDescription = ""; // Description field
                            string dbSupportPeriod2 = ""; // Compare support period field
                            string dbLanguage = ""; // Language field
                            string dbLanguageOther = ""; // Language other field
                            string dbSpeakProficient = ""; // Speak proficient field
                            string dbNDIS = ""; // NDIS field
                            DateTime dbCreatedOn = new DateTime(); // Created on field

                            string dbStAwaitGovt = ""; // Await govt status update field
                            string dbStCareOrder = ""; // Child care status update field
                            string dbStCaseMgt = ""; // Case mgt plan status update field
                            string dbStCaseMgtGoal = ""; // Case mgt plan goal status update field
                            string dbStCaseMgtReason = ""; // Case mgt plan reason status update field
                            string dbStCaseMgtReasonOth = ""; // Case mgt plan reason other status update field
                            string dbStEntryDate = ""; // Entry date status update field
                            string dbStFtPt = ""; // FT PT status update field
                            string dbStHomeless = ""; // Homeless status update field
                            string dbStIncome = ""; // Income status update field
                            string dbStLabourForce = ""; // Labour force status update field
                            string dbStLivingArrange = ""; // Living Arrange status update field
                            string dbStOccupancy = ""; // Occupancy status update field
                            string dbStResidential = ""; // Residential status update field
                            string dbStStudInd = ""; // Student indicator status update field
                            string dbStStudType = ""; // Student type status update field
                            string dbStSupportPeriod = ""; // Support period status update field
                            string dbStTenure = ""; // Tenure status update field
                            string dbStResubmit = ""; // Resubmit status update field
                            string dbStOngoing = ""; // Ongoing status update field

                            string dbSvSupportPeriod = ""; // Support period services field
                            string dbSvEntryDate = ""; // Entry date services field
                            string dbSvDescription = ""; // Description services field
                            string dbServArrange = ""; // Service arranged field
                            string dbServNeed = ""; // Service needed field
                            string dbServProvide = ""; // Service provided field

                            int varValidDate = 0; // Check for overlapping accom
                            string varLatestDate = ""; // Check last accom processed date
                            string dbAcSupportPeriod = ""; // Support period accom field
                            string dbAccomType = ""; // Accom type field
                            string dbAcLocation = ""; // Accom location field
                            string dbAcDateFrom = ""; // Accom start date field
                            string dbAcDateTo = ""; // Accom end date field

                            string dbFinSupportPeriod = ""; // Support period financial field
                            string dbFinEntryDate = ""; // Entry date financial field
                            string dbFinAmount = ""; // Financial amount field
                            string dbFinShor = ""; // Financial SHS field

                            string dbStepAgencyId = ""; // Turnaway agency field
                            string dbStepDate = ""; // Turnaway date field
                            string dbStepDob = ""; // Turnaway DOB field
                            string dbStepEstimated = ""; // Turnaway Estimated field
                            string dbStepFirstName = ""; // Turnaway First name field
                            string dbStepFirstServ = ""; // Turnaway first service field
                            string dbStepGender = ""; // Turnaway gender field
                            string dbStepIsPalm = ""; // Turnaway is palm client field
                            string dbStepIsStep = ""; // Turnaway PUH Id field
                            string dbStepName = ""; // Turnaway name field
                            string dbStepRelationship = ""; // Turnaway relationship field
                            string dbStepRelOther = ""; // Turnaway relationship (other) field
                            string dbStepSurname = ""; // Turnaway surname field
                            string dbStepTurnAssist = ""; // Turnaway assistance field
                            string dbStepTurnReason = ""; // Turnaway reason field
                            string dbStepTurnReasOth = ""; // Turnaway reason (other) field
                            string dbStepTurnRequest = ""; // Turnaway requested field
                            string dbStepTurnUrg = ""; // Turnaway urgency field
                            string dbPalmStepId = ""; // Turnaway recod id field
                            string dbStepIsStep2 = ""; // Turnaway PUH Id field comparison
                            string dbPalmStepId2 = ""; // Turnaway recod id field comparison
                            string dbStepRelationship2 = ""; // Turnaway relationship field comparison
                            string dbPalmClientSupportIdOld = ""; //Support Period id - old
                            string dbPuhIdOld = ""; //Presenting unit head old
                            string cleanFlagDate = ""; 
                            string debugDob = "";

                            string dbClient2 = ""; // Client number PUH
                            string dbDescription2 = ""; // Description PUH
                            string dbPuhId2 = ""; // PUH id comparison
                            string dbPuhRship2 = ""; // PUH relationship
                            string dbPalmClientSupportId2 = ""; // Support period PUH

                            // Variables for SU drop down list data
                            string varType = "";
                            string varDesc = "";
                            string varSHOR = "";

                            //Variable for setting old ids for imported records.
                            string varPrintSpId = "";
                            string varPrintPuhId = "";


                            int varDoSupport = 0; // Whether status update exists
                            int varServDone = 0; // Whether service provided
                            int varSubmission = 0; // Submission indicator
                            string varServStartDate = ""; // Service start date
                            string varServEndDate = ""; // Service end date

                            string varSurname = ""; // Surname for alpha code
                            string varFirstName = ""; // First name for alpha code
                            int varAge = 0; // Client age
                            int varPUHcount = 0; // Number in presenting unit
                            string varInPuh = "none"; // String for PUH Id
                            string varDobFlag = ""; // Dob flag

                            bool varDoEnd = false; // Whether accom ended
                            string varSuppId = ""; // Supp id for extract (no dashes)
                            string varPuhId = ""; // PUH id for extract (no dashes)
                            string varStepId = ""; // Unassist id for extract (no dashes)

                            string varServNeedM = "**"; // Services needed string
                            string varServProvideM = "**"; // Services provided string
                            string varServArrangeM = "**"; // Services arranged string

                            bool varAccomST = false; // Whether short term accom provided
                            bool varAccomMT = false; // Whether medium term accom provided
                            bool varAccomLT = false; // Whether long term accom provided

                            // Counters for header data
                            int varHHcount = 0;
                            int varHMcount = 0;
                            int varHYcount = 0;
                            int varINcount = 0;
                            int varREcount = 0;
                            int varSPcount = 0;
                            int varCPcount = 0;
                            int varSScount = 0;
                            int varAPcount = 0;
                            int varFScount = 0;
                            int varDFcount = 0;
                            int varTRcount = 0;
                            int varTScount = 0;
                            int varTAcount = 0;

                            // Used to get option set values
                            OptionSetValueCollection ovc = new OptionSetValueCollection();

                            // Get the fetch XML data and place in entity collection objects
                            EntityCollection result = _service.RetrieveMultiple(new FetchExpression(dbSHORList));
                            EntityCollection result2 = _service.RetrieveMultiple(new FetchExpression(dbDropList));
                            EntityCollection result4 = _service.RetrieveMultiple(new FetchExpression(dbStatusList));
                            EntityCollection result5 = _service.RetrieveMultiple(new FetchExpression(dbPrevList));
                            EntityCollection result7 = _service.RetrieveMultiple(new FetchExpression(dbServicesList));
                            EntityCollection result8 = _service.RetrieveMultiple(new FetchExpression(dbAccomList));
                            EntityCollection result9 = _service.RetrieveMultiple(new FetchExpression(dbFinancialList));
                            EntityCollection result10 = _service.RetrieveMultiple(new FetchExpression(dbStepList));
                            EntityCollection result11 = _service.RetrieveMultiple(new FetchExpression(dbTurnPuhList));

                            // Headers for client list
                            sbClientList.AppendLine("Type,Dynamics Id,Client Id,First Name,Surname,Start Date,");


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

                                //Reset the variables
                                varDoSupport = 0;
                                varServDone = 0;
                                sbSHORHomeSt.Length = 0;
                                sbSHORHomeMonth.Length = 0;
                                sbSHORHomeYear.Length = 0;
                                sbSHORFac.Length = 0;
                                sbSHORReas.Length = 0;
                                sbSHORAccomServ.Length = 0;
                                sbSHORServices.Length = 0;

                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (c.FormattedValues.Contains("new_address"))
                                    dbClient = c.FormattedValues["new_address"];
                                else if (c.Attributes.Contains("new_address"))
                                    dbClient = c.Attributes["new_address"].ToString();
                                else
                                    dbClient = "";

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

                                if (c.FormattedValues.Contains("new_dobestd"))
                                    dbDobEstD = c.FormattedValues["new_dobestd"];
                                else if (c.Attributes.Contains("new_dobestd"))
                                    dbDobEstD = c.Attributes["new_dobestd"].ToString();
                                else
                                    dbDobEstD = "";

                                if (c.FormattedValues.Contains("new_dobestm"))
                                    dbDobEstM = c.FormattedValues["new_dobestm"];
                                else if (c.Attributes.Contains("new_dobestm"))
                                    dbDobEstM = c.Attributes["new_dobestm"].ToString();
                                else
                                    dbDobEstM = "";

                                if (c.FormattedValues.Contains("new_dobesty"))
                                    dbDobEstY = c.FormattedValues["new_dobesty"];
                                else if (c.Attributes.Contains("new_dobesty"))
                                    dbDobEstY = c.Attributes["new_dobesty"].ToString();
                                else
                                    dbDobEstY = "";

                                // Set default dob flag
                                varDobFlag = "AAA";

                                if (string.IsNullOrEmpty(dbDobEstD) == false || string.IsNullOrEmpty(dbDobEstM) == false || string.IsNullOrEmpty(dbDobEstY) == false)
                                {


                                    // Insert day, month or year estimated part if not null (U or E)
                                    if (string.IsNullOrEmpty(dbDobEstD) == false)
                                    {
                                        varDobFlag = dbDobEstD.Substring(0, 1) + varDobFlag.Substring(1, 2);
                                    }
                                    if (string.IsNullOrEmpty(dbDobEstM) == false)
                                    {
                                        varDobFlag = varDobFlag.Substring(0, 1) + dbDobEstM.Substring(0, 1) + varDobFlag.Substring(2, 1);
                                    }

                                    if (string.IsNullOrEmpty(dbDobEstY) == false)
                                    {
                                        varDobFlag = varDobFlag.Substring(0, 2) + dbDobEstY.Substring(0, 1);
                                    }
                                }
                                else if (dbDobEst == "Yes")
                                {
                                    // Old format - If estimated is yes then do estimated for all
                                    varDobFlag = "EEE";
                                }

                                // Get new dog flag format
                                dbDobEst = varDobFlag;

                                if (c.FormattedValues.Contains("new_yeararrival"))
                                    dbYearArrival = c.FormattedValues["new_yeararrival"];
                                else if (c.Attributes.Contains("new_yeararrival"))
                                    dbYearArrival = c.Attributes["new_yeararrival"].ToString();
                                else
                                    dbYearArrival = "";

                                // Make sure numeric only
                                dbYearArrival = cleanString(dbYearArrival, "number");

                                if (c.FormattedValues.Contains("new_indigenous"))
                                    dbIndigenous = c.FormattedValues["new_indigenous"];
                                else if (c.Attributes.Contains("new_indigenous"))
                                    dbIndigenous = c.Attributes["new_indigenous"].ToString();
                                else
                                    dbIndigenous = "";

                                if (c.FormattedValues.Contains("new_palmddlcountry4.new_code"))
                                    dbCountry = c.FormattedValues["new_palmddlcountry4.new_code"];
                                else if (c.Attributes.Contains("new_palmddlcountry4.new_code"))
                                    dbCountry = c.Attributes["new_palmddlcountry4.new_code"].ToString();
                                else
                                    dbCountry = "";

                                if (c.FormattedValues.Contains("new_consent"))
                                    dbConsent = c.FormattedValues["new_consent"];
                                else if (c.Attributes.Contains("new_consent"))
                                    dbConsent = c.Attributes["new_consent"].ToString();
                                else
                                    dbConsent = "";

                                if (c.FormattedValues.Contains("new_palmddllanguage5.new_code"))
                                    dbLanguage = c.FormattedValues["new_palmddllanguage5.new_code"];
                                else if (c.Attributes.Contains("new_palmddllanguage5.new_code"))
                                    dbLanguage = c.Attributes["new_palmddllanguage5.new_code"].ToString();
                                else
                                    dbLanguage = "";

                                if (c.FormattedValues.Contains("new_palmddllanguage6.new_code"))
                                    dbLanguageOther = c.FormattedValues["new_palmddllanguage6.new_code"];
                                else if (c.Attributes.Contains("new_palmddllanguage6.new_code"))
                                    dbLanguageOther = c.Attributes["new_palmddllanguage6.new_code"].ToString();
                                else
                                    dbLanguageOther = "";

                                // Should use language other instead of language if one exists
                                if (String.IsNullOrEmpty(dbLanguageOther) == false)
                                    dbLanguage = dbLanguageOther;

                                if (c.FormattedValues.Contains("new_speakproficient"))
                                    dbSpeakProficient = c.FormattedValues["new_speakproficient"];
                                else if (c.Attributes.Contains("new_speakproficient"))
                                    dbSpeakProficient = c.Attributes["new_speakproficient"].ToString();
                                else
                                    dbSpeakProficient = "";

                                if (c.FormattedValues.Contains("new_consent"))
                                    dbConsent = c.FormattedValues["new_consent"];
                                else if (c.Attributes.Contains("new_consent"))
                                    dbConsent = c.Attributes["new_consent"].ToString();
                                else
                                    dbConsent = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_adf"))
                                    dbADF = c.FormattedValues["new_palmclientsupport1.new_adf"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_adf"))
                                    dbADF = c.Attributes["new_palmclientsupport1.new_adf"].ToString();
                                else
                                    dbADF = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_assessed"))
                                    dbAssessed = c.FormattedValues["new_palmclientsupport1.new_assessed"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_assessed"))
                                    dbAssessed = c.Attributes["new_palmclientsupport1.new_assessed"].ToString();
                                else
                                    dbAssessed = "";

                                // Convert date from American format to Australian format
                                dbAssessed = cleanDateAM(dbAssessed);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_awaitgovtpres"))
                                    dbAwaitGovtPres = c.FormattedValues["new_palmclientsupport1.new_awaitgovtpres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_awaitgovtpres"))
                                    dbAwaitGovtPres = c.Attributes["new_palmclientsupport1.new_awaitgovtpres"].ToString();
                                else
                                    dbAwaitGovtPres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_awaitgovtwkbef"))
                                    dbAwaitGovtWkBef = c.FormattedValues["new_palmclientsupport1.new_awaitgovtwkbef"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_awaitgovtwkbef"))
                                    dbAwaitGovtWkBef = c.Attributes["new_palmclientsupport1.new_awaitgovtwkbef"].ToString();
                                else
                                    dbAwaitGovtWkBef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_carepres"))
                                    dbCarePres = c.FormattedValues["new_palmclientsupport1.new_carepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_carepres"))
                                    dbCarePres = c.Attributes["new_palmclientsupport1.new_carepres"].ToString();
                                else
                                    dbCarePres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_carewkbef"))
                                    dbCareWkBef = c.FormattedValues["new_palmclientsupport1.new_carewkbef"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_carewkbef"))
                                    dbCareWkBef = c.Attributes["new_palmclientsupport1.new_carewkbef"].ToString();
                                else
                                    dbCareWkBef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_cessation"))
                                    dbCessation = c.FormattedValues["new_palmclientsupport1.new_cessation"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_cessation"))
                                    dbCessation = c.Attributes["new_palmclientsupport1.new_cessation"].ToString();
                                else
                                    dbCessation = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_diagnosedmh"))
                                    dbDiagnosedMH = c.FormattedValues["new_palmclientsupport1.new_diagnosedmh"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_diagnosedmh"))
                                    dbDiagnosedMH = c.Attributes["new_palmclientsupport1.new_diagnosedmh"].ToString();
                                else
                                    dbDiagnosedMH = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_discomm"))
                                    dbDisComm = c.FormattedValues["new_palmclientsupport1.new_discomm"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_discomm"))
                                    dbDisComm = c.Attributes["new_palmclientsupport1.new_discomm"].ToString();
                                else
                                    dbDisComm = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_dismob"))
                                    dbDisMob = c.FormattedValues["new_palmclientsupport1.new_dismob"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_dismob"))
                                    dbDisMob = c.Attributes["new_palmclientsupport1.new_dismob"].ToString();
                                else
                                    dbDisMob = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_disself"))
                                    dbDisSelf = c.FormattedValues["new_palmclientsupport1.new_disself"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_disself"))
                                    dbDisSelf = c.Attributes["new_palmclientsupport1.new_disself"].ToString();
                                else
                                    dbDisSelf = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_educationpres"))
                                    dbEducationPres = c.FormattedValues["new_palmclientsupport1.new_educationpres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_educationpres"))
                                    dbEducationPres = c.Attributes["new_palmclientsupport1.new_educationpres"].ToString();
                                else
                                    dbEducationPres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_facilities"))
                                    dbFacilities = c.FormattedValues["new_palmclientsupport1.new_facilities"]; //multi
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_facilities"))
                                    dbFacilities = c.Attributes["new_palmclientsupport1.new_facilities"].ToString();
                                else
                                    dbFacilities = "";

                                // Wrap asteriskes around facilities for better value matching
                                dbFacilities = getMult(dbFacilities);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_ftptpres"))
                                    dbFtPtPres = c.FormattedValues["new_palmclientsupport1.new_ftptpres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_ftptpres"))
                                    dbFtPtPres = c.Attributes["new_palmclientsupport1.new_ftptpres"].ToString();
                                else
                                    dbFtPtPres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_ftptwkbef"))
                                    dbFtPtWkBef = c.FormattedValues["new_palmclientsupport1.new_ftptwkbef"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_ftptwkbef"))
                                    dbFtPtWkBef = c.Attributes["new_palmclientsupport1.new_ftptwkbef"].ToString();
                                else
                                    dbFtPtWkBef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_homemonth"))
                                    dbHomeMonth = c.FormattedValues["new_palmclientsupport1.new_homemonth"]; //multi
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_homemonth"))
                                    dbHomeMonth = c.Attributes["new_palmclientsupport1.new_homemonth"].ToString();
                                else
                                    dbHomeMonth = "";

                                // Wrap asteriskes around homeless month for better value matching
                                dbHomeMonth = getMult(dbHomeMonth);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_homeyear"))
                                    dbHomeYear = c.FormattedValues["new_palmclientsupport1.new_homeyear"]; //multi
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_homeyear"))
                                    dbHomeYear = c.Attributes["new_palmclientsupport1.new_homeyear"].ToString();
                                else
                                    dbHomeYear = "";

                                // Wrap asteriskes around homeless year for better value matching
                                dbHomeYear = getMult(dbHomeYear);

                                //Support period id from new records created in dynamics - links to related records (fin, accom etc)
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_palmclientsupportid"))
                                    dbPalmClientSupportId = c.FormattedValues["new_palmclientsupport1.new_palmclientsupportid"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_palmclientsupportid"))
                                    dbPalmClientSupportId = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_palmclientsupportid").Value.ToString();
                                else
                                    dbPalmClientSupportId = "";

                                //Support period from id imported records.
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supportperiodidold"))
                                    dbPalmClientSupportIdOld = c.FormattedValues["new_palmclientsupport1.new_supportperiodidold"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supportperiodidold"))
                                    dbPalmClientSupportIdOld = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_supportperiodidold").Value.ToString();
                                else
                                    dbPalmClientSupportIdOld = "";

                                //Fix as imported PALM support period ids do not include palm client id.
                                //This might be the biggest bodge in the whole system.
                                if (String.IsNullOrEmpty(dbPalmClientSupportIdOld) == false)
                                {
                                    if(dbPalmClientSupportIdOld.Length < 5) //Support period id's from SHIP are 7 characters - doubtful a palm client would have had more than 10000 support periods so roughly works.
                                    {
                                        varPrintSpId = dbClient + "_" + dbPalmClientSupportIdOld; 
                                    }
                                    else
                                    {
                                        varPrintSpId = dbPalmClientSupportIdOld;
                                    }
                                    
                                }
                                else
                                {
                                    varPrintSpId = dbPalmClientSupportId;
                                }



                                // Remove dashes from support period id for extract
                                varSuppId = "";
                                if (String.IsNullOrEmpty(varPrintSpId) == false)
                                    varSuppId = varPrintSpId.Replace("-", "");

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_incomepres"))
                                    dbIncomePres = c.FormattedValues["new_palmclientsupport1.new_incomepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_incomepres"))
                                    dbIncomePres = c.Attributes["new_palmclientsupport1.new_incomepres"].ToString();
                                else
                                    dbIncomePres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_incomewkbef"))
                                    dbIncomeWkBef = c.FormattedValues["new_palmclientsupport1.new_incomewkbef"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_incomewkbef"))
                                    dbIncomeWkBef = c.Attributes["new_palmclientsupport1.new_incomewkbef"].ToString();
                                else
                                    dbIncomeWkBef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_labourforcepres"))
                                    dbLabourForcePres = c.FormattedValues["new_palmclientsupport1.new_labourforcepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_labourforcepres"))
                                    dbLabourForcePres = c.Attributes["new_palmclientsupport1.new_labourforcepres"].ToString();
                                else
                                    dbLabourForcePres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_labourforcewkbef"))
                                    dbLabourForceWkBef = c.FormattedValues["new_palmclientsupport1.new_labourforcewkbef"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_labourforcewkbef"))
                                    dbLabourForceWkBef = c.Attributes["new_palmclientsupport1.new_labourforcewkbef"].ToString();
                                else
                                    dbLabourForceWkBef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_livingarrangepres"))
                                    dbLivingArrangePres = c.FormattedValues["new_palmclientsupport1.new_livingarrangepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_livingarrangepres"))
                                    dbLivingArrangePres = c.Attributes["new_palmclientsupport1.new_livingarrangepres"].ToString();
                                else
                                    dbLivingArrangePres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_livingarrangewkbef"))
                                    dbLivingArrangeWkBef = c.FormattedValues["new_palmclientsupport1.new_livingarrangewkbef"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_livingarrangewkbef"))
                                    dbLivingArrangeWkBef = c.Attributes["new_palmclientsupport1.new_livingarrangewkbef"].ToString();
                                else
                                    dbLivingArrangeWkBef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_mentalillnessinfo"))
                                    dbMentalIllnessInfo = c.FormattedValues["new_palmclientsupport1.new_mentalillnessinfo"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_mentalillnessinfo"))
                                    dbMentalIllnessInfo = c.Attributes["new_palmclientsupport1.new_mentalillnessinfo"].ToString();
                                else
                                    dbMentalIllnessInfo = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_mhservicesrecd"))
                                    dbMHServicesRecd = c.FormattedValues["new_palmclientsupport1.new_mhservicesrecd"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_mhservicesrecd"))
                                    dbMHServicesRecd = c.Attributes["new_palmclientsupport1.new_mhservicesrecd"].ToString();
                                else
                                    dbMHServicesRecd = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_newclient"))
                                    dbNewClient = c.FormattedValues["new_palmclientsupport1.new_newclient"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_newclient"))
                                    dbNewClient = c.Attributes["new_palmclientsupport1.new_newclient"].ToString();
                                else
                                    dbNewClient = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_occupancypres"))
                                    dbOccupancyPres = c.FormattedValues["new_palmclientsupport1.new_occupancypres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_occupancypres"))
                                    dbOccupancyPres = c.Attributes["new_palmclientsupport1.new_occupancypres"].ToString();
                                else
                                    dbOccupancyPres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_occupancywkbef"))
                                    dbOccupancyWkBef = c.FormattedValues["new_palmclientsupport1.new_occupancywkbef"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_occupancywkbef"))
                                    dbOccupancyWkBef = c.Attributes["new_palmclientsupport1.new_occupancywkbef"].ToString();
                                else
                                    dbOccupancyWkBef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_preslocality"))
                                    dbPresLocality = c.FormattedValues["new_palmclientsupport1.new_preslocality"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_preslocality"))
                                    dbPresLocality = c.Attributes["new_palmclientsupport1.new_preslocality"].ToString();
                                else
                                    dbPresLocality = "";

                                if (c.FormattedValues.Contains("new_palmddllocality2.new_state"))
                                    dbPresState = c.FormattedValues["new_palmddllocality2.new_state"];
                                else if (c.Attributes.Contains("new_palmddllocality2.new_state"))
                                    dbPresState = c.Attributes["new_palmddllocality2.new_state"].ToString();
                                else
                                    dbPresState = "";

                                if (c.FormattedValues.Contains("new_palmddllocality2.new_postcode"))
                                    dbPresPostcode = c.FormattedValues["new_palmddllocality2.new_postcode"];
                                else if (c.Attributes.Contains("new_palmddllocality2.new_postcode"))
                                    dbPresPostcode = c.Attributes["new_palmddllocality2.new_postcode"].ToString();
                                else
                                    dbPresPostcode = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_primreason"))
                                    dbPrimReason = c.FormattedValues["new_palmclientsupport1.new_primreason"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_primreason"))
                                    dbPrimReason = c.Attributes["new_palmclientsupport1.new_primreason"].ToString();
                                else
                                    dbPrimReason = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_program"))
                                    dbProgram = c.FormattedValues["new_palmclientsupport1.new_program"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_program"))
                                    dbProgram = c.Attributes["new_palmclientsupport1.new_program"].ToString();
                                else
                                    dbProgram = "";

                                // We need to get the entity id for the presenting unit field
                                if (c.Attributes.Contains("new_palmclientsupport1.new_puhid"))
                                {
                                    // Get the entity id for the presenting unit using the entity reference object
                                    getEntity = (EntityReference)c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_puhid").Value;
                                    dbPuhId = getEntity.Id.ToString();
                                }
                                else if (c.FormattedValues.Contains("new_palmclientsupport1.new_puhid"))
                                    dbPuhId = c.FormattedValues["new_palmclientsupport1.new_puhid"];
                                else
                                    dbPuhId = "";

                                /*
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_puhidold"))
                                    dbPuhIdOld = c.FormattedValues["new_palmclientsupport1.new_puhidold"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_puhidold"))
                                    dbPuhIdOld = c.Attributes["new_palmclientsupport1.new_puhidold"].ToString();
                                else
                                    dbPuhIdOld = "";
                                */

                                //Support period from id imported records.
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_puhidold"))
                                    dbPuhIdOld = c.FormattedValues["new_palmclientsupport1.new_puhidold"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_puhidold"))
                                    dbPuhIdOld = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_puhidold").Value.ToString();
                                else
                                    dbPuhIdOld = "";

                                if (String.IsNullOrEmpty(dbPuhIdOld) == false)
                                    varPrintPuhId = dbPuhIdOld;
                                else
                                    varPrintPuhId = dbPuhId;




                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_puhrship"))
                                    dbPuhRship = c.FormattedValues["new_palmclientsupport1.new_puhrship"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_puhrship"))
                                    dbPuhRship = c.Attributes["new_palmclientsupport1.new_puhrship"].ToString();
                                else
                                    dbPuhRship = "";

                                // If relationship is self, the PUHID should be self
                                if (dbPuhRship.ToLower() == "self")
                                    dbPuhId = varPrintSpId; //dbPalmClientSupportId

                                 /*
                                 if (string.IsNullOrEmpty(dbPuhIdOld) == false)
                                     varPrintPuhId = dbPuhIdOld;
                                 else
                                     varPrintPuhId = dbPuhId;
                                 */

                                 // Remove dashes from presenting unit id for extract
                                 varPuhId = "";
                                if (String.IsNullOrEmpty(varPrintPuhId) == false)
                                    varPuhId = varPrintPuhId.Replace("-", "");

                                /*if (c.FormattedValues.Contains("new_palmclientsupport1.new_puhrshipother"))
                                    dbPuhRshipOther = c.FormattedValues["new_palmclientsupport1.new_puhrshipother"];
                                else*/ if (c.Attributes.Contains("new_palmclientsupport1.new_puhrshipother"))
                                    //dbPuhRshipOther = c.Attributes["new_palmclientsupport1.new_puhrshipother"].ToString();
                                dbPuhRshipOther = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_puhrshipother").Value.ToString();
                                else
                                    dbPuhRshipOther = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_reasons"))
                                    dbReasons = c.FormattedValues["new_palmclientsupport1.new_reasons"]; // multi
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_reasons"))
                                    dbReasons = c.Attributes["new_palmclientsupport1.new_reasons"].ToString();
                                else
                                    dbReasons = "";

                                // Wrap asteriskes around reasons for better value matching
                                dbReasons = getMult(dbReasons);

                                //C07.021.02 | Main reason for assistance must be in reasons for presenting
                                if (dbReasons.IndexOf("*" + dbPrimReason + "*") == -1)
                                       sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Main reason for assistance is not in list of reasons for presenting");

                                //if (c.FormattedValues.Contains("new_palmclientsupport1.new_reasonsoth")) - CHANGE
                                //dbReasonsOth = c.FormattedValues["new_palmclientsupport1.new_reasonsoth"];
                                /*else*/
                                if (c.Attributes.Contains("new_palmclientsupport1.new_reasonsoth"))
                                    //dbReasonsOth = c.Attributes["new_palmclientsupport1.new_reasonsoth"].ToString();
                                    dbReasonsOth = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_reasonsoth").Value.ToString();
                                else
                                    dbReasonsOth = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_residentialpres"))
                                    dbResidentialPres = c.FormattedValues["new_palmclientsupport1.new_residentialpres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_residentialpres"))
                                    dbResidentialPres = c.Attributes["new_palmclientsupport1.new_residentialpres"].ToString();
                                else
                                    dbResidentialPres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_residentialwkbef"))
                                    dbResidentialWkBef = c.FormattedValues["new_palmclientsupport1.new_residentialwkbef"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_residentialwkbef"))
                                    dbResidentialWkBef = c.Attributes["new_palmclientsupport1.new_residentialwkbef"].ToString();
                                else
                                    dbResidentialWkBef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_shordate"))
                                    dbShorDate = c.FormattedValues["new_palmclientsupport1.new_shordate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_shordate"))
                                    dbShorDate = c.Attributes["new_palmclientsupport1.new_shordate"].ToString();
                                else
                                    dbShorDate = "";

                                // Convert date from American format to Australian format
                                dbShorDate = cleanDateAM(dbShorDate);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_sourceref"))
                                    dbSourceRef = c.FormattedValues["new_palmclientsupport1.new_sourceref"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_sourceref"))
                                    dbSourceRef = c.Attributes["new_palmclientsupport1.new_sourceref"].ToString();
                                else
                                    dbSourceRef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_startdate"))
                                    dbStartDate = c.FormattedValues["new_palmclientsupport1.new_startdate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_startdate"))
                                    dbStartDate = c.Attributes["new_palmclientsupport1.new_startdate"].ToString();
                                else
                                    dbStartDate = "";

                                // Convert date from American format to Australian format
                                dbStartDate = cleanDateAM(dbStartDate);

                                //Fix issue if assessed date not entered for a client's support period.
                                if (String.IsNullOrEmpty(dbAssessed) == true)
                                    dbAssessed = dbStartDate;

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_enddate"))
                                    dbEndDate = c.FormattedValues["new_palmclientsupport1.new_enddate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_enddate"))
                                    dbEndDate = c.Attributes["new_palmclientsupport1.new_enddate"].ToString();
                                else
                                    dbEndDate = "";

                                // Convert date from American format to Australian format
                                dbEndDate = cleanDateAM(dbEndDate);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_studindpres"))
                                    dbStudIndPres = c.FormattedValues["new_palmclientsupport1.new_studindpres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_studindpres"))
                                    dbStudIndPres = c.Attributes["new_palmclientsupport1.new_studindpres"].ToString();
                                else
                                    dbStudIndPres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_studindwkbef"))
                                    dbStudIndWkBef = c.FormattedValues["new_palmclientsupport1.new_studindwkbef"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_studindwkbef"))
                                    dbStudIndWkBef = c.Attributes["new_palmclientsupport1.new_studindwkbef"].ToString();
                                else
                                    dbStudIndWkBef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_studtypepres"))
                                    dbStudTypePres = c.FormattedValues["new_palmclientsupport1.new_studtypepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_studtypepres"))
                                    dbStudTypePres = c.Attributes["new_palmclientsupport1.new_studtypepres"].ToString();
                                else
                                    dbStudTypePres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_studtypewkbef"))
                                    dbStudTypeWkBef = c.FormattedValues["new_palmclientsupport1.new_studtypewkbef"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_studtypewkbef"))
                                    dbStudTypeWkBef = c.Attributes["new_palmclientsupport1.new_studtypewkbef"].ToString();
                                else
                                    dbStudTypeWkBef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_tenurepres"))
                                    dbTenurePres = c.FormattedValues["new_palmclientsupport1.new_tenurepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_tenurepres"))
                                    dbTenurePres = c.Attributes["new_palmclientsupport1.new_tenurepres"].ToString();
                                else
                                    dbTenurePres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_tenurewkbef"))
                                    dbTenureWkBef = c.FormattedValues["new_palmclientsupport1.new_tenurewkbef"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_tenurewkbef"))
                                    dbTenureWkBef = c.Attributes["new_palmclientsupport1.new_tenurewkbef"].ToString();
                                else
                                    dbTenureWkBef = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_timeperm"))
                                    dbTimePerm = c.FormattedValues["new_palmclientsupport1.new_timeperm"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_timeperm"))
                                    dbTimePerm = c.Attributes["new_palmclientsupport1.new_timeperm"].ToString();
                                else
                                    dbTimePerm = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.ownerid"))
                                    dbOwnerId = c.FormattedValues["new_palmclientsupport1.ownerid"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_ownerid"))
                                    dbOwnerId = c.Attributes["new_palmclientsupport1.new_ownerid"].ToString();
                                else
                                    dbOwnerId = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_wkbeflocality"))
                                    dbWkBefLocality = c.FormattedValues["new_palmclientsupport1.new_wkbeflocality"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_wkbeflocality"))
                                    dbWkBefLocality = c.Attributes["new_palmclientsupport1.new_wkbeflocality"].ToString();
                                else
                                    dbWkBefLocality = "";

                                if (c.FormattedValues.Contains("new_palmddllocality3.new_state"))
                                    dbWkBefState = c.FormattedValues["new_palmddllocality3.new_state"];
                                else if (c.Attributes.Contains("new_palmddllocality3.new_state"))
                                    dbWkBefState = c.Attributes["new_palmddllocality3.new_state"].ToString();
                                else
                                    dbWkBefState = "";

                                if (c.FormattedValues.Contains("new_palmddllocality3.new_postcode"))
                                    dbWkBefPostcode = c.FormattedValues["new_palmddllocality3.new_postcode"];
                                else if (c.Attributes.Contains("new_palmddllocality3.new_postcode"))
                                    dbWkBefPostcode = c.Attributes["new_palmddllocality3.new_postcode"].ToString();
                                else
                                    dbWkBefPostcode = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_description"))
                                    dbDescription = c.FormattedValues["new_palmclientsupport1.new_description"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_description"))
                                    dbDescription = c.Attributes["new_palmclientsupport1.new_description"].ToString();
                                else
                                    dbDescription = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_ndis"))
                                    dbNDIS = c.FormattedValues["new_palmclientsupport1.new_ndis"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_ndis"))
                                    dbNDIS = c.Attributes["new_palmclientsupport1.new_ndis"].ToString();
                                else
                                    dbNDIS = "";

                                // Get created on date
                                if (c.FormattedValues.Contains("createdon"))
                                    DateTime.TryParse(c.FormattedValues["createdon"], out dbCreatedOn);
                                else if (c.Attributes.Contains("createdon"))
                                    DateTime.TryParse(c.Attributes["createdon"].ToString(), out dbCreatedOn);

                                // Loop through status update data
                                foreach (var s in result4.Entities)
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

                                    // We need to get the entity id for the support period field for comparisons
                                    if (s.Attributes.Contains("new_supportperiod"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)s.Attributes["new_supportperiod"];
                                        dbStSupportPeriod = getEntity.Id.ToString();
                                    }
                                    else if (s.FormattedValues.Contains("new_supportperiod"))
                                        dbStSupportPeriod = s.FormattedValues["new_supportperiod"];
                                    else
                                        dbStSupportPeriod = "";

                                    // Need to see if same support period
                                    if (dbStSupportPeriod == dbPalmClientSupportId)
                                    {
                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (s.FormattedValues.Contains("new_awaitgovt"))
                                            dbStAwaitGovt = s.FormattedValues["new_awaitgovt"];
                                        else if (s.Attributes.Contains("new_awaitgovt"))
                                            dbStAwaitGovt = s.Attributes["new_awaitgovt"].ToString();
                                        else
                                            dbStAwaitGovt = "";

                                        if (s.FormattedValues.Contains("new_careorder"))
                                            dbStCareOrder = s.FormattedValues["new_careorder"];
                                        else if (s.Attributes.Contains("new_careorder"))
                                            dbStCareOrder = s.Attributes["new_careorder"].ToString();
                                        else
                                            dbStCareOrder = "";

                                        if (s.FormattedValues.Contains("new_casemgt"))
                                            dbStCaseMgt = s.FormattedValues["new_casemgt"];
                                        else if (s.Attributes.Contains("new_casemgt"))
                                            dbStCaseMgt = s.Attributes["new_casemgt"].ToString();
                                        else
                                            dbStCaseMgt = "";

                                        if (s.FormattedValues.Contains("new_casemgtgoal"))
                                            dbStCaseMgtGoal = s.FormattedValues["new_casemgtgoal"];
                                        else if (s.Attributes.Contains("new_casemgtgoal"))
                                            dbStCaseMgtGoal = s.Attributes["new_casemgtgoal"].ToString();
                                        else
                                            dbStCaseMgtGoal = "";

                                        if (s.FormattedValues.Contains("new_casemgtreason"))
                                            dbStCaseMgtReason = s.FormattedValues["new_casemgtreason"];
                                        else if (s.Attributes.Contains("new_casemgtreason"))
                                            dbStCaseMgtReason = s.Attributes["new_casemgtreason"].ToString();
                                        else
                                            dbStCaseMgtReason = "";

                                        if (s.FormattedValues.Contains("new_casemgtreasonoth"))
                                            dbStCaseMgtReasonOth = s.FormattedValues["new_casemgtreasonoth"];
                                        else if (s.Attributes.Contains("new_casemgtreasonoth"))
                                            dbStCaseMgtReasonOth = s.Attributes["new_casemgtreasonoth"].ToString();
                                        else
                                            dbStCaseMgtReasonOth = "";

                                        if (s.FormattedValues.Contains("new_entrydate"))
                                            dbStEntryDate = s.FormattedValues["new_entrydate"];
                                        else if (s.Attributes.Contains("new_entrydate"))
                                            dbStEntryDate = s.Attributes["new_entrydate"].ToString();
                                        else
                                            dbStEntryDate = "";

                                        // Convert date from American format to Australian format
                                        dbStEntryDate = cleanDateAM(dbStEntryDate);

                                        if (s.FormattedValues.Contains("new_ftpt"))
                                            dbStFtPt = s.FormattedValues["new_ftpt"];
                                        else if (s.Attributes.Contains("new_ftpt"))
                                            dbStFtPt = s.Attributes["new_ftpt"].ToString();
                                        else
                                            dbStFtPt = "";

                                        if (s.FormattedValues.Contains("new_homeless"))
                                            dbStHomeless = s.FormattedValues["new_homeless"];
                                        else if (s.Attributes.Contains("new_homeless"))
                                            dbStHomeless = s.Attributes["new_homeless"].ToString();
                                        else
                                            dbStHomeless = "";

                                        // Wrap asteriskes around facilities for better value matching
                                        dbStHomeless = getMult(dbStHomeless);

                                        if (s.FormattedValues.Contains("new_income"))
                                            dbStIncome = s.FormattedValues["new_income"];
                                        else if (s.Attributes.Contains("new_income"))
                                            dbStIncome = s.Attributes["new_income"].ToString();
                                        else
                                            dbStIncome = "";

                                        if (s.FormattedValues.Contains("new_labourforce"))
                                            dbStLabourForce = s.FormattedValues["new_labourforce"];
                                        else if (s.Attributes.Contains("new_labourforce"))
                                            dbStLabourForce = s.Attributes["new_labourforce"].ToString();
                                        else
                                            dbStLabourForce = "";

                                        if (s.FormattedValues.Contains("new_livingarrange"))
                                            dbStLivingArrange = s.FormattedValues["new_livingarrange"];
                                        else if (s.Attributes.Contains("new_livingarrange"))
                                            dbStLivingArrange = s.Attributes["new_livingarrange"].ToString();
                                        else
                                            dbStLivingArrange = "";

                                        if (s.FormattedValues.Contains("new_occupancy"))
                                            dbStOccupancy = s.FormattedValues["new_occupancy"];
                                        else if (s.Attributes.Contains("new_occupancy"))
                                            dbStOccupancy = s.Attributes["new_occupancy"].ToString();
                                        else
                                            dbStOccupancy = "";

                                        if (s.FormattedValues.Contains("new_residential"))
                                            dbStResidential = s.FormattedValues["new_residential"];
                                        else if (s.Attributes.Contains("new_residential"))
                                            dbStResidential = s.Attributes["new_residential"].ToString();
                                        else
                                            dbStResidential = "";

                                        if (s.FormattedValues.Contains("new_studind"))
                                            dbStStudInd = s.FormattedValues["new_studind"];
                                        else if (s.Attributes.Contains("new_studind"))
                                            dbStStudInd = s.Attributes["new_studind"].ToString();
                                        else
                                            dbStStudInd = "";

                                        if (s.FormattedValues.Contains("new_studtype"))
                                            dbStStudType = s.FormattedValues["new_studtype"];
                                        else if (s.Attributes.Contains("new_studtype"))
                                            dbStStudType = s.Attributes["new_studtype"].ToString();
                                        else
                                            dbStStudType = "";

                                        if (s.FormattedValues.Contains("new_tenure"))
                                            dbStTenure = s.FormattedValues["new_tenure"];
                                        else if (s.Attributes.Contains("new_tenure"))
                                            dbStTenure = s.Attributes["new_tenure"].ToString();
                                        else
                                            dbStTenure = "";

                                        /*
                                        if (s.FormattedValues.Contains("new_resubmit"))
                                            dbStResubmit = s.FormattedValues["new_resubmit"];
                                        else if (s.Attributes.Contains("new_resubmit"))
                                            dbStResubmit = s.Attributes["new_resubmit"].ToString();
                                        else
                                            dbStResubmit = "";
                                        */
                                        //Fix for removing option to not resubmit support period information.
                                        dbStResubmit = "Yes";

                                        if (s.FormattedValues.Contains("new_ongoing"))
                                            dbStOngoing = s.FormattedValues["new_ongoing"];
                                        else if (s.Attributes.Contains("new_ongoing"))
                                            dbStOngoing = s.Attributes["new_ongoing"].ToString();
                                        else
                                            dbStOngoing = "";

                                        varDoSupport = 1; //Status update done
                                        break;
                                    } // Same support period
                                } //Status update loop

                                //[SHOR](23) | Submission Value
                                //0 = no previous period
                                //1 = previous period but closed
                                //2 = previous period but ongoing

                                //Loop through previous status updates to see if this client had one in the previous month
                                varSubmission = 0;

                                // Previous status loop
                                foreach (var p in result5.Entities)
                                {
                                    // We need to get the entity id for the support period field for comparisons
                                    if (p.Attributes.Contains("new_supportperiod"))
                                    {
                                        // Get the entity id for the support period using the entity reference object
                                        getEntity = (EntityReference)p.Attributes["new_supportperiod"];
                                        dbSupportPeriod2 = getEntity.Id.ToString();
                                    }
                                    else if (p.FormattedValues.Contains("new_supportperiod"))
                                        dbSupportPeriod2 = p.FormattedValues["new_supportperiod"];
                                    else
                                        dbSupportPeriod2 = "";

                                    // If the support period matches then the submission indicator is 2 (ongoing)
                                    if (String.IsNullOrEmpty(dbPalmClientSupportId) == false && dbPalmClientSupportId == dbSupportPeriod2)
                                    {
                                        varSubmission = 1;
                                        break;
                                    }
                                }

                                //Allow a resubmit
                                if (varSubmission == 2 && dbStResubmit == "Yes")
                                    varSubmission = 1;

                                //Palm Fix for February - no previous status update before 1 Feb 2013
                                //If month is Feb or less and year, then set 2 to 0
                                if (varSubmission == 2 && varMonthStart <= Convert.ToDateTime("1-Feb-2013"))
                                    varSubmission = 0;

                                //Palm broken
                                if (String.IsNullOrEmpty(dbStartDate) == false)
                                {
                                    if (Convert.ToDateTime(dbStartDate) < Convert.ToDateTime("1-Jul-2015"))
                                        varSubmission = 1; //Was 2
                                }

                                // Check for last perm locality
                                if (String.IsNullOrEmpty(dbPresLocality) == false && String.IsNullOrEmpty(dbPresPostcode) == false && String.IsNullOrEmpty(dbPresState) == false)
                                {
                                    //Postcode needs to be 4 digits
                                    if (dbPresPostcode.Length == 1)
                                        dbPresPostcode = "000" + dbPresPostcode;
                                    else if (dbPresPostcode.Length == 2)
                                        dbPresPostcode = "00" + dbPresPostcode;
                                    else if (dbPresPostcode.Length == 3)
                                        dbPresPostcode = "0" + dbPresPostcode;
                                    else if (dbPresPostcode.Length > 4)
                                        dbPresPostcode = "9999";
                                }
                                else
                                {
                                    // Set to don't know
                                    dbPresLocality = "";
                                    dbPresPostcode = "";
                                    dbPresState = "Don't Know";
                                }

                                // Check for week before locality
                                if (String.IsNullOrEmpty(dbWkBefLocality) == false && String.IsNullOrEmpty(dbWkBefPostcode) == false && String.IsNullOrEmpty(dbWkBefState) == false)
                                {
                                    //Postcode needs to be 4 digits
                                    if (dbWkBefPostcode.Length == 1)
                                        dbWkBefPostcode = "000" + dbWkBefPostcode;
                                    else if (dbWkBefPostcode.Length == 2)
                                        dbWkBefPostcode = "00" + dbWkBefPostcode;
                                    else if (dbWkBefPostcode.Length == 3)
                                        dbWkBefPostcode = "0" + dbWkBefPostcode;
                                    else if (dbWkBefPostcode.Length > 4)
                                        dbWkBefPostcode = "9999";
                                }
                                else
                                {
                                    // Set to don't know
                                    dbWkBefLocality = "";
                                    dbWkBefPostcode = "";
                                    dbWkBefState = "Don't Know";
                                }

                                //Get default values if there is no data
                                //[SHOR](76) | Insitutions
                                if (String.IsNullOrEmpty(dbFacilities) == true)
                                    dbFacilities = "*99*";

                                //[SHOR](79) | Reasons
                                if (String.IsNullOrEmpty(dbReasons) == true)
                                    dbReasons = "*99*";

                                //[SHOR](83) | Homeless History last month
                                if (String.IsNullOrEmpty(dbHomeMonth) == true)
                                    dbHomeMonth = "*99*";

                                //[SHOR](86) | Homeless History last year
                                if (String.IsNullOrEmpty(dbHomeYear) == true)
                                    dbHomeYear = "*99*";

                                if (String.IsNullOrEmpty(dbStHomeless) == true)
                                    dbStHomeless = "*99*";

                                //Ensure there are no conflicting values
                                //C08.003.02 | Error if institution is no institution and has other values
                                if (dbFacilities.IndexOf("*88*") > -1 || dbFacilities.IndexOf("No institution") > -1)
                                    dbFacilities = "*88*";

                                //C08.003.03 | Error if institution is don't know and has other values
                                if (dbFacilities.IndexOf("*99*") > -1 || dbFacilities.IndexOf("Don't know") > -1)
                                    dbFacilities = "*99*";

                                //C09.003.02 | Error if reasons is don't know and has other values
                                if (dbReasons.IndexOf("*99*") > -1 || dbReasons.IndexOf("Don't know") > -1)
                                    dbReasons = "*99*";

                                //C10.003.02 | Error if homeless month is not homeless and there are other values
                                if (dbHomeMonth.IndexOf("*3*") > -1)
                                    dbHomeMonth = "*3*";

                                //C10.003.03 | Error if homeless month is dont know and there are other values
                                if (dbHomeMonth.IndexOf("*99*") > -1)
                                    dbHomeMonth = "*99*";

                                //C11.003.02 | Error if homeless year is not homeless and there are other values
                                if (dbHomeYear.IndexOf("*3*") > -1)
                                    dbHomeYear = "*3*";

                                //C11.003.03 | Error if homeless year is dont know and there are other values
                                if (dbHomeYear.IndexOf("*99*") > -1)
                                    dbHomeYear = "*99*";

                                //C16.003.02 | Error if homeless status update is not homeless and there are other values
                                if (dbStHomeless.IndexOf("*3*") > -1)
                                    dbStHomeless = "*3*";

                                //C16.003.03 | Error if homeless status update is dont know and there are other values
                                if (dbStHomeless.IndexOf("*99*") > -1)
                                    dbStHomeless = "*99*";

                                //[SHOR](80) | Reasons Other
                                if (String.IsNullOrEmpty(dbReasonsOth) == true)
                                    dbReasonsOth = "Not Stated";
                                else if (dbReasonsOth.Length > 100)
                                    dbReasonsOth = "Not Stated";


                                //[SHOR](32) | Consent Indicator
                                if (dbConsent == "No")
                                {
                                    dbConsent = "2";
                                    dbFacilities = "*0*";
                                }
                                else
                                    dbConsent = "1";

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

                                    if (d.FormattedValues.Contains("new_shor"))
                                        varSHOR = d.FormattedValues["new_shor"];
                                    else if (d.Attributes.Contains("new_shor"))
                                        varSHOR = d.Attributes["new_shor"].ToString();
                                    else
                                        varSHOR = "";

                                    // Make sure SHOR value is numeric, or set to 0
                                    varSHOR = cleanString(varSHOR, "number");
                                    if (String.IsNullOrEmpty(varSHOR) == true)
                                        varSHOR = "0";

                                    // If the drop down type is ongoing, compare with ongoing field
                                    // If the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "ongoing" && varDesc == dbStOngoing)
                                    {
                                        dbStOngoing = varSHOR + "";
                                        break;
                                    }
                                }

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

                                    if (d.FormattedValues.Contains("new_shor"))
                                        varSHOR = d.FormattedValues["new_shor"];
                                    else if (d.Attributes.Contains("new_shor"))
                                        varSHOR = d.Attributes["new_shor"].ToString();
                                    else
                                        varSHOR = "";

                                    // Make sure SHOR value is numeric, or set to 0
                                    varSHOR = cleanString(varSHOR, "number");
                                    if (String.IsNullOrEmpty(varSHOR) == true)
                                        varSHOR = "0";

                                    // If the type is indigenous and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "indigenous" && varDesc == dbIndigenous)
                                        dbIndigenous = varSHOR;

                                    // If the type is puhrship and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "puhrship" && varDesc == dbPuhRship)
                                        dbPuhRship = varSHOR;

                                    // If the type is sourceref and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "sourceref" && varDesc == dbSourceRef)
                                        dbSourceRef = varSHOR;

                                    // If the type is reasons and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "reasons" && varDesc == dbPrimReason)
                                        dbPrimReason = varSHOR;

                                    // If the type is state and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "state" && varDesc == dbPresState)
                                        dbPresState = varSHOR;

                                    // If the type is timeperm and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "timeperm" && varDesc == dbTimePerm)
                                        dbTimePerm = varSHOR;

                                    // If the type is state and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "state" && varDesc == dbWkBefState)
                                        dbWkBefState = varSHOR;

                                    // If the type is studenttype and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "studenttype" && varDesc == dbStudTypeWkBef)
                                        dbStudTypeWkBef = varSHOR;

                                    // If the type is studenttype and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "studenttype" && varDesc == dbStudTypePres)
                                        dbStudTypePres = varSHOR;

                                    // If the type is education and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "education" && varDesc == dbEducationPres)
                                        dbEducationPres = varSHOR;

                                    // If the type is income and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "income" && varDesc == dbIncomeWkBef)
                                        dbIncomeWkBef = varSHOR;

                                    // If the type is income and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "income" && varDesc == dbIncomePres)
                                        dbIncomePres = varSHOR;

                                    // If the type is livingarrange and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "livingarrange" && varDesc == dbLivingArrangeWkBef)
                                        dbLivingArrangeWkBef = varSHOR;

                                    // If the type is livingarrange and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "livingarrange" && varDesc == dbLivingArrangePres)
                                        dbLivingArrangePres = varSHOR;

                                    // If the type is labour and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "labour" && varDesc == dbLabourForceWkBef)
                                        dbLabourForceWkBef = varSHOR;

                                    // If the type is labour and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "labour" && varDesc == dbLabourForcePres)
                                        dbLabourForcePres = varSHOR;

                                    // If the type is residential and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "residential" && varDesc == dbResidentialWkBef)
                                        dbResidentialWkBef = varSHOR;

                                    // If the type is residential and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "residential" && varDesc == dbResidentialPres)
                                        dbResidentialPres = varSHOR;

                                    // If the type is tenure and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "tenure" && varDesc == dbTenureWkBef)
                                        dbTenureWkBef = varSHOR;

                                    // If the type is tenure and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "tenure" && varDesc == dbTenurePres)
                                        dbTenurePres = varSHOR;

                                    // If the type is occupancy and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "occupancy" && varDesc == dbOccupancyWkBef)
                                        dbOccupancyWkBef = varSHOR;

                                    // If the type is occupancy and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "occupancy" && varDesc == dbOccupancyPres)
                                        dbOccupancyPres = varSHOR;

                                    // If the type is childcare and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "childcare" && varDesc == dbCareWkBef)
                                        dbCareWkBef = varSHOR;

                                    // If the type is childcare and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "childcare" && varDesc == dbCarePres)
                                        dbCarePres = varSHOR;

                                    // If the type is mhservices and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "mhservices" && varDesc == dbMHServicesRecd)
                                        dbMHServicesRecd = varSHOR;

                                    // If the type is mentalillness and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "mentalillness" && varDesc == dbMentalIllnessInfo)
                                        dbMentalIllnessInfo = varSHOR;

                                    // If the type is labour and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "labour" && varDesc == dbStLabourForce)
                                        dbStLabourForce = varSHOR;

                                    // If the type is studenttype and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "studenttype" && varDesc == dbStStudType)
                                        dbStStudType = varSHOR;

                                    // If the type is livingarrange and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "livingarrange" && varDesc == dbStLivingArrange)
                                        dbStLivingArrange = varSHOR;

                                    // If the type is occupancy and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "occupancy" && varDesc == dbStOccupancy)
                                        dbStOccupancy = varSHOR;

                                    // If the type is residential and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "residential" && varDesc == dbStResidential)
                                        dbStResidential = varSHOR;

                                    // If the type is tenure and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "tenure" && varDesc == dbStTenure)
                                        dbStTenure = varSHOR;

                                    // If the type is income and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "income" && varDesc == dbStIncome)
                                        dbStIncome = varSHOR;

                                    // If the type is childcare and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "childcare" && varDesc == dbStCareOrder)
                                        dbStCareOrder = varSHOR;

                                    // If the type is casereason and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "casereason" && varDesc == dbStCaseMgtReason)
                                        dbStCaseMgtReason = varSHOR;

                                    // If the type is cessation and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "cessation" && varDesc == dbCessation)
                                        dbCessation = varSHOR + "";

                                    // If the type is shordis and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "shordis" && varDesc == dbDisComm)
                                        dbDisComm = varSHOR + "";

                                    // If the type is shordis and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "shordis" && varDesc == dbDisMob)
                                        dbDisMob = varSHOR + "";

                                    // If the type is shordis and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "shordis" && varDesc == dbDisSelf)
                                        dbDisSelf = varSHOR + "";

                                    // If the type is speak proficient and the string values match, set the numeric value to the SHOR value from the Drop Down List
                                    if (varType == "speakproficient" && varDesc == dbSpeakProficient)
                                    {
                                        dbSpeakProficient = varSHOR + "";

                                        //Error checking for speak proficient.

                                        if (Convert.ToDateTime(dbStartDate) > Convert.ToDateTime("1-Jul-2019"))
                                        {
                                            if (dbLanguageOther != "1201")
                                            {
                                                if (Convert.ToDateTime(dbDob) < DateTime.Now.AddYears(-5))
                                                {
                                                    if (dbSpeakProficient == "0") //"NA"
                                                    {
                                                        sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - If client 'Language other than English Spoken at Home' is not english, client is 5 or more years old, and 'Proficiency in spoken English' must not be Not Applicable.");
                                                    }
                                                }
                                                else
                                                {
                                                    if (dbSpeakProficient != "0") //Not "NA"
                                                    {
                                                        sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - If client is less than 5 years old, 'Proficiency in spoken English' must be Not Applicable.");
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (dbSpeakProficient != "0") //Not "NA"
                                                {
                                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - If client 'Language other than English Spoken at Home' is english, 'Proficiency in spoken English' must be Not Applicable.");
                                                }
                                            }

                                        }
                                        else
                                        {
                                            if (dbSpeakProficient != "0") //Not "NA"
                                            {
                                                sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - If support period started before 1-July-2019, 'Proficiency in spoken English' must be Not Applicable.");
                                            }
                                        }
                                    }
                                        

                                    


                                    // Create the homeless parts of the extract
                                    if (varType == "homeless")
                                    {
                                        //Create the homeless status update part of the extract
                                        if (String.IsNullOrEmpty(dbStHomeless) == false)
                                        {
                                            // If the homeless variable contains the number or description, and the ongoing indicator is not ended last month
                                            if ((dbStHomeless.IndexOf("*" + varDesc + "*") > -1 || dbStHomeless.IndexOf("*" + varSHOR + "*") > -1) && dbStOngoing != "3")
                                            {
                                                // Homeless Status Update Count
                                                varHHcount++;
                                                // Append to status update homeless
                                                sbSHORHomeSt.AppendLine("            <SP_CP_Homeless_History>");
                                                sbSHORHomeSt.AppendLine("               <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                sbSHORHomeSt.AppendLine("               <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                sbSHORHomeSt.AppendLine("               <Previously_Homeless_Ind_CP_End>" + varSHOR + "</Previously_Homeless_Ind_CP_End>");
                                                sbSHORHomeSt.AppendLine("            </SP_CP_Homeless_History>");
                                            }
                                        }

                                        //Create the homeless month part of the extract
                                        if (String.IsNullOrEmpty(dbHomeMonth) == false && varSubmission != 2)
                                        {
                                            // If the homeless variable contains the number or description, and the ongoing indicator is not ended last month
                                            if (dbHomeMonth.IndexOf("*" + varDesc + "*") > -1 || dbHomeMonth.IndexOf("*" + varSHOR + "*") > -1)
                                            {
                                                //Homeless Month Count
                                                varHMcount++;
                                                // Append to homeless month
                                                sbSHORHomeMonth.AppendLine("         <SP_Homeless_History_Mth>");
                                                sbSHORHomeMonth.AppendLine("            <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                sbSHORHomeMonth.AppendLine("            <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                sbSHORHomeMonth.AppendLine("            <Previously_Homeless_Ind>" + varSHOR + "</Previously_Homeless_Ind>");
                                                sbSHORHomeMonth.AppendLine("         </SP_Homeless_History_Mth>");
                                            }
                                        }

                                        //Create the homeless year part of the extract
                                        if (String.IsNullOrEmpty(dbHomeYear) == false && varSubmission != 2)
                                        {
                                            // If the homeless variable contains the number or description, and the ongoing indicator is not ended last month
                                            if (dbHomeYear.IndexOf("*" + varDesc + "*") > -1 || dbHomeYear.IndexOf("*" + varSHOR + "*") > -1)
                                            {
                                                //Homeless Month Count
                                                varHYcount++;
                                                // Append to homeless year
                                                sbSHORHomeYear.AppendLine("         <SP_Homeless_History_Yr>");
                                                sbSHORHomeYear.AppendLine("            <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                sbSHORHomeYear.AppendLine("            <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                sbSHORHomeYear.AppendLine("            <Previously_Homeless_Ind>" + varSHOR + "</Previously_Homeless_Ind>");
                                                sbSHORHomeYear.AppendLine("         </SP_Homeless_History_Yr>");
                                            }
                                        }

                                    } //Homeless

                                    //Create the facilities part of the extract
                                    if (varType == "facilities" && varSubmission != 2)
                                    {
                                        if (String.IsNullOrEmpty(dbFacilities) == false)
                                        {
                                            // If the facilities variable contains the number or description
                                            if (dbFacilities.IndexOf("*" + varDesc + "*") > -1 || dbFacilities.IndexOf("*" + varSHOR + "*") > -1)
                                            {
                                                //Institutions Count
                                                varINcount++;
                                                // Append to facilities
                                                sbSHORFac.AppendLine("         <SP_Institutions>");
                                                sbSHORFac.AppendLine("            <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                sbSHORFac.AppendLine("            <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                sbSHORFac.AppendLine("            <Facility_Type_Recently_Left>" + varSHOR + "</Facility_Type_Recently_Left>");
                                                sbSHORFac.AppendLine("         </SP_Institutions>");
                                            }
                                        }

                                    } // Facilities

                                    //Create the reasons part of the extract
                                    if (varType == "reasons" && varSubmission != 2)
                                    {
                                        if (String.IsNullOrEmpty(dbReasons) == false)
                                        {
                                            // If the reasons variable contains the number or description
                                            if (dbReasons.IndexOf("*" + varDesc + "*") > -1 || dbReasons.IndexOf("*" + varSHOR + "*") > -1)
                                            {
                                                //Reasons Count
                                                varREcount++;
                                                // Append to reasons
                                                sbSHORReas.AppendLine("         <SP_Reasons>");
                                                sbSHORReas.AppendLine("            <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                sbSHORReas.AppendLine("            <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                sbSHORReas.AppendLine("            <Assist_Reason_Present>" + varSHOR + "</Assist_Reason_Present>");

                                                if (varDesc == "Other")
                                                    sbSHORReas.AppendLine("            <Assist_Reason_Present_Other>" + cleanString(dbReasonsOth, "normal") + "</Assist_Reason_Present_Other>");
                                                else
                                                    sbSHORReas.AppendLine("            <Assist_Reason_Present_Other xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");

                                                sbSHORReas.AppendLine("         </SP_Reasons>");
                                            }
                                        }

                                    } // Reasons

                                } //Dop down List

                                // Set defaults for don't know or not applicable
                                if (dbCountry == "Not Applicable")
                                    dbCountry = "9999";
                                else if (dbCountry == "Don't know")
                                    dbCountry = "0000";

                                if (dbLanguage == "Don't know")
                                    dbLanguage = "0002";
                                else if (dbLanguage == "Not Applicable")
                                    dbLanguage = "0002";
                                else if (dbLanguage == "2")
                                    dbLanguage = "0002";

                                //[SHOR](21) | Support Period ID
                                if (String.IsNullOrEmpty(dbClient) == true)
                                {
                                    dbClient = "0";
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Must have client number");
                                }
                                else if (dbClient.Length > 50)
                                {
                                    dbClient = "0";
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client number greater than 50 characters");
                                }

                                //[Feb Fix]
                                //New client number format was created 1 Feb 2013. Earlier start dates were reset to 1 February to avoid errors
                                if (String.IsNullOrEmpty(dbStartDate) == false)
                                {
                                    if (Convert.ToDateTime(dbStartDate) < Convert.ToDateTime("1-Feb-2013"))
                                    {
                                        dbStartDate = "1-Feb-2013";
                                        dbAssessed = "1-Feb-2013";
                                    }
                                }

                                //[SHOR](22)(109) | Episode Start Date cannot be earlier than 1895 or empty
                                if (String.IsNullOrEmpty(dbStartDate) == false)
                                {
                                    if (Convert.ToDateTime(dbStartDate).Year < 1895)
                                        sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Start Date");
                                }
                                else
                                    sbErrorList.AppendLine("Data Missing: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Start Date");

                                //[SHOR](24) | Letters of Family Name
                                // Append characters and extract 2nd, 3rd and 5th character
                                varSurname = dbSurname.ToUpper() + "22222";
                                varSurname = cleanString(varSurname, "slk");
                                varSurname = varSurname.Substring(1, 1) + varSurname.Substring(2, 1) + varSurname.Substring(4, 1);
                                // If empty set to 999
                                if (varSurname == "222")
                                    varSurname = "999";

                                
                                //[SHOR](25) | Letters of Given Name
                                // Append characters and extract 2nd, 3rd character
                                varFirstName = dbFirstName.ToUpper() + "22222";
                                varFirstName = cleanString(varFirstName, "slk");

                                varFirstName = varFirstName.Substring(1, 1) + varFirstName.Substring(2, 1);
                                // If empty set to 99
                                if (varFirstName == "22")
                                    varFirstName = "99";

                                

                                //[SHOR](26) | Sex
                                if (dbGender == "Male")
                                    dbGender = "1";
                                else if (dbGender == "Other")
                                {
                                    if (Convert.ToDateTime(dbStartDate) < Convert.ToDateTime("1-Jul-2019") || varMonthStart < Convert.ToDateTime("1-Jul-2019"))
                                        dbGender = "2";
                                    else
                                        dbGender = "3";
                                }
                                else
                                    dbGender = "2";

                                //[SHOR](28) | Date of Birth Estimate
                                if (dbDobEst == "estimated")
                                    dbDobEst = "EEE";
                                else if (dbDobEst == "not estimated")
                                    dbDobEst = "AAA";

                                // Make sure valid length
                                dbDobEst = dbDobEst + "AAA";
                                dbDobEst = dbDobEst.Substring(0, 3);

                                // Remove illegal characters
                                if (dbDobEst.Substring(0, 1) != "A" && dbDobEst.Substring(0, 1) != "E" && dbDobEst.Substring(0, 1) != "U")
                                    dbDobEst = "AAA";
                                if (dbDobEst.Substring(1, 1) != "A" && dbDobEst.Substring(1, 1) != "E" && dbDobEst.Substring(1, 1) != "U")
                                    dbDobEst = "AAA";
                                if (dbDobEst.Substring(2, 1) != "A" && dbDobEst.Substring(2, 1) != "E" && dbDobEst.Substring(2, 1) != "U")
                                    dbDobEst = "AAA";

                                //[SHOR](27) | Date of Birth
                                if (String.IsNullOrEmpty(dbDob) == false)
                                {
                                    // Year must be greater or equal to 1880 and client cannot be greater than 116 years old
                                    if (Convert.ToDateTime(dbDob).Year >= 1880)
                                    {
                                        if (Convert.ToDateTime(dbDob) < Convert.ToDateTime(DateTime.Now.AddYears(-116)))
                                        {
                                            dbDob = "1-Jan-1880";
                                            dbDobEst = "UUU";
                                        }
                                    }
                                    else
                                    {
                                        // Default for invalid date
                                        dbDob = "1-Jan-1880";
                                        dbDobEst = "UUU";
                                        sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - DOB");
                                    }
                                }
                                else
                                {
                                    // Default for invalid date
                                    dbDob = "1-Jan-1880";
                                    dbDobEst = "UUU";
                                }

                                // Default for invalid date
                                if (dbDobEst == "UUU")
                                    dbDob = "1-Jan-1880";

                                // Calculate the client age
                                varAge = varMonthEnd.Year - Convert.ToDateTime(dbDob).Year;
                                if (varMonthEnd.Month < Convert.ToDateTime(dbDob).Month || (varMonthEnd.Month == Convert.ToDateTime(dbDob).Month && varMonthEnd.Day < Convert.ToDateTime(dbDob).Day))
                                    varAge = varAge - 1;

                                //[SHOR](29) | Year of Arrival In Australia
                                if (!Int32.TryParse(dbYearArrival, out varCheckInt))
                                    dbYearArrival = "9999";
                                else if (dbYearArrival == "0")
                                    dbYearArrival = "0000";

                                //[SHOR](30) | Indigenous Status
                                if (!Int32.TryParse(dbIndigenous, out varCheckInt))
                                    dbIndigenous = "99";

                                //[SHOR](31) | Country of Birth
                                if (!Int32.TryParse(dbCountry, out varCheckInt))
                                    dbCountry = "0000";

                                //[SHOR](33) | Assistance Date
                                if (String.IsNullOrEmpty(dbAssessed) == false)
                                {
                                    // Year cannot be greater than 1880
                                    if (Convert.ToDateTime(dbAssessed).Year < 1880)
                                    {
                                        dbAssessed = "";
                                        sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Assessed Date");
                                    }
                                }
                                else
                                {
                                    // Alert of missing date
                                    dbAssessed = "";
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Assessed Date");
                                }

                                //[SHOR](34) | New Client
                                if (dbNewClient == "Yes")
                                    dbNewClient = "1";
                                else if (dbNewClient == "No")
                                    dbNewClient = "2";
                                else
                                    dbNewClient = "3";

                                // Create SLK
                                //if (String.IsNullOrEmpty(dbAssessed) == false && String.IsNullOrEmpty(dbDob) == false)
                                //    varSLK = varSurname + varFirstName + cleanDateS(Convert.ToDateTime(dbDob)) + dbGender + "_" + cleanDateSLK(Convert.ToDateTime(dbCreatedOn));
                                //else
                                //    varSLK = "Error";

                                //[SHOR](36) | Relationship to Presenting Unit Head
                                if (!Int32.TryParse(dbPuhRship, out varCheckInt))
                                    dbPuhRship = "99";

                                //[SHOR](37) | Relationship to Presenting Unit Head Other
                                if (dbPuhRship == "12" || dbPuhRship == "15")
                                {
                                    // Other must have value if PUH Rhsip = 12 or 15
                                    if (String.IsNullOrEmpty(dbPuhRshipOther) == true)
                                        dbPuhRshipOther = "No Value";
                                    else if (dbPuhRshipOther.Length > 100)
                                        dbPuhRshipOther = "No Value";
                                }
                                else
                                    dbPuhRshipOther = ""; // Set to empty if not 12 or 15

                                //[SHOR](38) | Count in presenting unit
                                varPUHcount = 0;
                                varInPuh = "none";

                                // Loop through support periods
                                foreach (var c2 in result.Entities)
                                {
                                    //varTest = "STARTING ATTRIBUTES:\r\n";

                                    //foreach (KeyValuePair<String, Object> attribute in c2.Attributes)
                                    //{
                                    //    varTest += (attribute.Key + ": " + attribute.Value + "\r\n");
                                    //}

                                    //varTest += "STARTING FORMATTED:\r\n";

                                    //foreach (KeyValuePair<String, String> value in c2.FormattedValues)
                                    //{
                                    //    varTest += (value.Key + ": " + value.Value + "\r\n");
                                    //}

                                    // Get support period id
                                    if (c2.FormattedValues.Contains("new_palmclientsupport1.new_palmclientsupportid"))
                                        dbPalmClientSupportId2 = c2.FormattedValues["new_palmclientsupport1.new_palmclientsupportid"];
                                    else if (c2.Attributes.Contains("new_palmclientsupport1.new_palmclientsupportid"))
                                        dbPalmClientSupportId2 = c2.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_palmclientsupportid").Value.ToString();
                                    else
                                        dbPalmClientSupportId2 = "";

                                    // Get puhid id
                                    if (c2.Attributes.Contains("new_palmclientsupport1.new_puhid"))
                                    {
                                        getEntity = (EntityReference)c2.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_puhid").Value;
                                        dbPuhId2 = getEntity.Id.ToString();
                                    }
                                    else if (c2.FormattedValues.Contains("new_palmclientsupport1.new_puhid"))
                                        dbPuhId2 = c2.FormattedValues["new_palmclientsupport1.new_puhid"];
                                    else
                                        dbPuhId2 = "";

                                    // Get relationship to PUH
                                    if (c2.FormattedValues.Contains("new_palmclientsupport1.new_puhrship"))
                                        dbPuhRship2 = c2.FormattedValues["new_palmclientsupport1.new_puhrship"];
                                    else if (c2.Attributes.Contains("new_palmclientsupport1.new_puhrship"))
                                        dbPuhRship2 = c2.Attributes["new_palmclientsupport1.new_puhrship"].ToString();
                                    else
                                        dbPuhRship2 = "";

                                    // Set puhid to support id if relationship self
                                    if (dbPuhRship2.ToLower() == "self")
                                        dbPuhId2 = dbPalmClientSupportId2;

                                    // Add to the PUH count if this person's puhid is the same
                                    if (dbPuhId == dbPuhId2)
                                        varPUHcount++;

                                    //Get the PUH ID with the support period ID
                                    if (dbPuhId2 == dbPuhId && varInPuh == "none")
                                        varInPuh = dbPuhId2;

                                } //puh Loop

                                //Add 1 to PUH count for null and self
                                //if (String.IsNullOrEmpty(dbPuhId) == true && dbPuhRship.ToLower() == "self")
                                //{
                                //    varPUHcount++;
                                //    dbPuhId = dbPalmClientSupportId;
                                //}

                                //[SHOR](35) | Presenting Unit Head ID
                                if (String.IsNullOrEmpty(dbPuhId) == true)
                                {
                                    // Set default to self if empty
                                    dbPuhId = dbPalmClientSupportId;
                                    dbPuhRshipOther = "";
                                    dbPuhRship = "1";
                                    //sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - No presenting unit head id");
                                }
                                else if (dbPuhId.Length > 50)
                                {
                                    // Must be less than 50 characters
                                    dbPuhId = "0";
                                    dbPuhRshipOther = "";
                                    dbPuhRship = "1";
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Presenting unit head ID greater than 50 characters");
                                }

                                //[SHOR](39) | Formal referral source
                                if (!Int32.TryParse(dbSourceRef, out varCheckInt))
                                    dbSourceRef = "99";

                                // Can no longer use value 22
                                if (dbSourceRef == "22" && Convert.ToDateTime(dbStartDate) < Convert.ToDateTime("1-Jul-2019"))
                                    dbSourceRef = "19";

                                //[SHOR](40) | Main reason for assistance
                                if (!Int32.TryParse(dbPrimReason, out varCheckInt))
                                    dbPrimReason = "99";

                                //[SHOR](41) | Most recent locality
                                if (String.IsNullOrEmpty(dbPresLocality) == false && dbPresLocality.Contains("Unknown") && dbPresLocality.Length == 7)
                                    dbPresLocality = cleanString(dbPresLocality, "name");
                                else if (String.IsNullOrEmpty(dbPresLocality) == false && dbPresLocality.Contains("Unknown") && dbPresLocality.Length != 7) // May need to be changed in future if township called "unknown" is created
                                    dbPresLocality = "Unknown";


                                //[SHOR](42) | Most recent Postcode
                                if (!Int32.TryParse(dbPresPostcode, out varCheckInt))
                                    dbPresPostcode = "";

                                //[SHOR](43) | Most recent state
                                if (!Int32.TryParse(dbPresState, out varCheckInt))
                                    dbPresState = "99";

                                //[SHOR](44) | Time since most recent address
                                if (!Int32.TryParse(dbTimePerm, out varCheckInt))
                                    dbTimePerm = "99";

                                //[SHOR](45) | Locality the week before
                                if (String.IsNullOrEmpty(dbWkBefLocality) == false && dbWkBefLocality.Contains("Unknown") && dbWkBefLocality.Length == 7)
                                    dbWkBefLocality = cleanString(dbWkBefLocality, "name");
                                else if (String.IsNullOrEmpty(dbWkBefLocality) == false && dbWkBefLocality.Contains("Unknown") && dbWkBefLocality.Length != 7) 
                                    dbWkBefLocality = "Unknown"; 

                                //[SHOR](46) | Postcode the week before
                                if (!Int32.TryParse(dbWkBefPostcode, out varCheckInt))
                                    dbWkBefPostcode = "";

                                //[SHOR](47) | State the week before
                                if (!Int32.TryParse(dbWkBefState, out varCheckInt))
                                    dbWkBefState = "99";

                                //ToDo: Make sure not hardcoded

                                //[SHOR](48) | Awaiting govenment benefit week before
                                if (dbAwaitGovtWkBef == "Yes")
                                    dbAwaitGovtWkBef = "1";
                                else if (dbAwaitGovtWkBef == "No")
                                    dbAwaitGovtWkBef = "2";
                                else if (dbAwaitGovtWkBef == "Don't know")
                                    dbAwaitGovtWkBef = "99";
                                else if (dbAwaitGovtWkBef == "Not applicable")
                                    dbAwaitGovtWkBef = "0";
                                else
                                    dbAwaitGovtWkBef = "99";

                                //[SHOR](49) | Awaiting govenment benefit presenting
                                if (dbAwaitGovtPres == "Yes")
                                    dbAwaitGovtPres = "1";
                                else if (dbAwaitGovtPres == "No")
                                    dbAwaitGovtPres = "2";
                                else if (dbAwaitGovtPres == "Don't know")
                                    dbAwaitGovtPres = "99";
                                else if (dbAwaitGovtPres == "Not applicable")
                                    dbAwaitGovtPres = "0";
                                else
                                    dbAwaitGovtPres = "99";

                                //[SHOR](50) | Student indicator week before
                                if (dbStudIndWkBef == "Yes")
                                    dbStudIndWkBef = "1";
                                else if (dbStudIndWkBef == "No")
                                    dbStudIndWkBef = "2";
                                else if (dbStudIndWkBef == "Don't know")
                                    dbStudIndWkBef = "99";
                                else
                                    dbStudIndWkBef = "99";

                                //[SHOR](51) | Student indicator presenting
                                if (dbStudIndPres == "Yes")
                                    dbStudIndPres = "1";
                                else if (dbStudIndPres == "No")
                                    dbStudIndPres = "2";
                                else if (dbStudIndPres == "Don't know")
                                    dbStudIndPres = "99";
                                else
                                    dbStudIndPres = "99";

                                //[SHOR](52) | Student type week before
                                if (!Int32.TryParse(dbStudTypeWkBef, out varCheckInt))
                                    dbStudTypeWkBef = "99";

                                //[SHOR](53) | Student type presenting
                                if (!Int32.TryParse(dbStudTypePres, out varCheckInt))
                                    dbStudTypePres = "99";

                                //[SHOR](54) | Education at present
                                if (!Int32.TryParse(dbEducationPres, out varCheckInt))
                                    dbEducationPres = "99";

                                //[SHOR](55) | Source of income week before
                                if (!Int32.TryParse(dbIncomeWkBef, out varCheckInt))
                                    dbIncomeWkBef = "99";

                                //[SHOR](56) | Source of income presenting
                                if (!Int32.TryParse(dbIncomePres, out varCheckInt))
                                    dbIncomePres = "99";

                                //[SHOR](57) | Living arrangements week before
                                if (!Int32.TryParse(dbLivingArrangeWkBef, out varCheckInt))
                                    dbLivingArrangeWkBef = "99";

                                //[SHOR](58) | Living arrangements presenting
                                if (!Int32.TryParse(dbLivingArrangePres, out varCheckInt))
                                    dbLivingArrangePres = "99";

                                //[SHOR](59) | Labour force status week before
                                if (!Int32.TryParse(dbLabourForceWkBef, out varCheckInt))
                                    dbLabourForceWkBef = "99";

                                //[SHOR](60) | Labour force status presenting
                                if (!Int32.TryParse(dbLabourForcePres, out varCheckInt))
                                    dbLabourForcePres = "99";

                                //[SHOR](61) | FT / PT status week before
                                if (dbFtPtWkBef == "Full-time")
                                    dbFtPtWkBef = "1";
                                else if (dbFtPtWkBef == "Part-time")
                                    dbFtPtWkBef = "2";
                                else if (dbFtPtWkBef == "Don't know")
                                    dbFtPtWkBef = "99";
                                else if (dbFtPtWkBef == "Not applicable")
                                    dbFtPtWkBef = "0";
                                else
                                    dbFtPtWkBef = "99";

                                //[SHOR](62) | FT / PT status presenting
                                if (dbFtPtPres == "Full-time")
                                    dbFtPtPres = "1";
                                else if (dbFtPtPres == "Part-time")
                                    dbFtPtPres = "2";
                                else if (dbFtPtPres == "Don't know")
                                    dbFtPtPres = "99";
                                else if (dbFtPtPres == "Not applicable")
                                    dbFtPtPres = "0";
                                else
                                    dbFtPtPres = "99";

                                //[SHOR](63) | Residential week before
                                if (!Int32.TryParse(dbResidentialWkBef, out varCheckInt))
                                    dbResidentialWkBef = "99";

                                //[SHOR](64) | Residential presenting
                                if (!Int32.TryParse(dbResidentialPres, out varCheckInt))
                                    dbResidentialPres = "99";

                                //[SHOR](65) | Tenure week before
                                if (!Int32.TryParse(dbTenureWkBef, out varCheckInt))
                                    dbTenureWkBef = "99";

                                //[SHOR](66) | Tenure presenting
                                if (!Int32.TryParse(dbTenurePres, out varCheckInt))
                                    dbTenurePres = "99";

                                //[SHOR](67) | Occupancy week before
                                if (!Int32.TryParse(dbOccupancyWkBef, out varCheckInt))
                                    dbOccupancyWkBef = "99";

                                //[SHOR](68) | Occupancy presenting
                                if (!Int32.TryParse(dbOccupancyPres, out varCheckInt))
                                    dbOccupancyPres = "99";

                                //[SHOR](69) | Child care week before
                                if (!Int32.TryParse(dbCareWkBef, out varCheckInt))
                                    dbCareWkBef = "99";

                                //[SHOR](70) | Child care presenting
                                if (!Int32.TryParse(dbCarePres, out varCheckInt))
                                    dbCarePres = "99";

                                //[SHOR](71) | Diagnosed mental health
                                if (dbDiagnosedMH == "Yes")
                                    dbDiagnosedMH = "1";
                                else if (dbDiagnosedMH == "No")
                                    dbDiagnosedMH = "2";
                                else if (dbDiagnosedMH == "Don't know")
                                    dbDiagnosedMH = "99";
                                else if (dbDiagnosedMH == "Not applicable")
                                    dbDiagnosedMH = "0";
                                else
                                    dbDiagnosedMH = "99";

                                //[SHOR](72) | Mental health services received
                                if (!Int32.TryParse(dbMHServicesRecd, out varCheckInt))
                                    dbMHServicesRecd = "99";

                                //[SHOR](73) | Source of mental illness information
                                if (!Int32.TryParse(dbMentalIllnessInfo, out varCheckInt))
                                    dbMentalIllnessInfo = "0";

                                //[SHOR](new) | Disability indicators
                                if (dbDisSelf != "1" && dbDisSelf != "2" && dbDisSelf != "3" && dbDisSelf != "4")
                                    dbDisSelf = "99";

                                if (dbDisMob != "1" && dbDisMob != "2" && dbDisMob != "3" && dbDisMob != "4")
                                    dbDisMob = "99";

                                if (dbDisComm != "1" && dbDisComm != "2" && dbDisComm != "3" && dbDisComm != "4")
                                    dbDisComm = "99";

                                // Make sure speak proficient numeric
                                if (!Int32.TryParse(dbSpeakProficient, out varCheckInt))
                                    dbSpeakProficient = "99";

                                //[SHOR](110)(111) | First day of contact / Last day of contact
                                // Reset values
                                varServStartDate = "";
                                varServEndDate = "";

                                varAccomST = false; // Whether short term accom provided
                                varAccomMT = false; // Whether medium term accom provided
                                varAccomLT = false; // Whether long term accom provided

                                // Loop through accom records to get service dates. The data will be added to the extract in a later section below
                                foreach (var s in result8.Entities)
                                {
                                    // We need to get the entity id for the support period field for comparisons
                                    if (s.Attributes.Contains("new_supportperiod"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)s.Attributes["new_supportperiod"];
                                        dbAcSupportPeriod = getEntity.Id.ToString();
                                    }
                                    else if (s.FormattedValues.Contains("new_supportperiod"))
                                        dbAcSupportPeriod = s.FormattedValues["new_supportperiod"];
                                    else
                                        dbAcSupportPeriod = "";

                                    // Need to see if same support period
                                    if (dbPalmClientSupportId == dbAcSupportPeriod)
                                    {
                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (s.FormattedValues.Contains("new_datefrom"))
                                            dbAcDateFrom = s.FormattedValues["new_datefrom"];
                                        else if (s.Attributes.Contains("new_datefrom"))
                                            dbAcDateFrom = s.Attributes["new_datefrom"].ToString();
                                        else
                                            dbAcDateFrom = "";

                                        // Convert date from American format to Australian format
                                        dbAcDateFrom = cleanDateAM(dbAcDateFrom);

                                        if (s.FormattedValues.Contains("new_dateto"))
                                            dbAcDateTo = s.FormattedValues["new_dateto"];
                                        else if (s.Attributes.Contains("new_dateto"))
                                            dbAcDateTo = s.Attributes["new_dateto"].ToString();
                                        else
                                            dbAcDateTo = "";

                                        // Convert date from American format to Australian format
                                        dbAcDateTo = cleanDateAM(dbAcDateTo);

                                        if (s.FormattedValues.Contains("new_location"))
                                            dbAcLocation = s.FormattedValues["new_location"];
                                        else if (s.Attributes.Contains("new_location"))
                                            dbAcLocation = s.Attributes["new_location"].ToString();
                                        else
                                            dbAcLocation = "";

                                        if (s.FormattedValues.Contains("new_accomtype"))
                                            dbAccomType = s.FormattedValues["new_accomtype"];
                                        else if (s.Attributes.Contains("new_accomtype"))
                                            dbAccomType = s.Attributes["new_accomtype"].ToString();
                                        else
                                            dbAccomType = "";

                                        // Process data if from date valid
                                        if (String.IsNullOrEmpty(dbAcDateFrom) == false)
                                        {
                                            /*
                                            // Old code when accom was stored in the services table. This check is done below. Uncomment if extract no longer works
                                            //If accom start date falls in the current month, get the service start and end dates
                                            if (Convert.ToDateTime(dbAcDateFrom) >= varMonthStart && Convert.ToDateTime(dbAcDateFrom) <= varMonthEnd)
                                            {
                                                //Get start date and end date
                                                if (String.IsNullOrEmpty(varServStartDate) == true)
                                                    varServStartDate = dbAcDateFrom;

                                                if (String.IsNullOrEmpty(varServEndDate) == true)
                                                    varServEndDate = dbAcDateFrom;
                                                else if (Convert.ToDateTime(varServEndDate) < Convert.ToDateTime(dbAcDateFrom))
                                                    varServEndDate = dbAcDateFrom;

                                            } //accomStart >= month start
                                            */

                                            //Reset valid date counter
                                            varValidDate = 0;
                                            // The accom has a start and end date
                                            if (String.IsNullOrEmpty(dbAcDateFrom) == false && String.IsNullOrEmpty(dbAcDateTo) == false)
                                            {
                                                // Get numeric value
                                                if (dbAccomType == "Short term or emergency accommodation")
                                                    dbAccomType = "1";
                                                else if (dbAccomType == "Medium term/transitional accommodation")
                                                    dbAccomType = "2";
                                                else if (dbAccomType == "Long term accommodation")
                                                    dbAccomType = "3";
                                                else
                                                    dbAccomType = "";

                                                // Check if accom overlaps current period
                                                if (Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthEnd) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateTo) <= Convert.ToDateTime(varMonthEnd))
                                                    varValidDate = 1;

                                                if (Convert.ToDateTime(dbAcDateFrom) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthEnd) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateTo) <= Convert.ToDateTime(varMonthEnd))
                                                    varValidDate = 1;

                                                if (Convert.ToDateTime(dbAcDateFrom) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthEnd) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthEnd))
                                                    varValidDate = 1;

                                                if (Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthEnd) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthEnd))
                                                    varValidDate = 1;

                                            }
                                            // The accom has only an end date
                                            else if (String.IsNullOrEmpty(dbAcDateFrom) == false)
                                            {
                                                // Get numeric value
                                                if (dbAccomType == "Short term or emergency accommodation")
                                                    dbAccomType = "1";
                                                else if (dbAccomType == "Medium term/transitional accommodation")
                                                    dbAccomType = "2";
                                                else if (dbAccomType == "Long term accommodation")
                                                    dbAccomType = "3";
                                                else
                                                    dbAccomType = "";

                                                // Check if accom overlaps current period
                                                if (Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthEnd))
                                                    varValidDate = 1;
                                            }

                                            //If accommodation exists and the date falls within this month, get the service start and end dates
                                            if (varValidDate == 1 && String.IsNullOrEmpty(dbAccomType) == false)
                                            {

                                                //Get start date and end date
                                                if (String.IsNullOrEmpty(varServStartDate) == true)
                                                    varServStartDate = dbAcDateFrom;
                                                else if (Convert.ToDateTime(dbAcDateFrom) < Convert.ToDateTime(varServStartDate))
                                                    varServStartDate = dbAcDateFrom;

                                                if (String.IsNullOrEmpty(varServEndDate) == true)
                                                    varServEndDate = dbAcDateFrom;
                                                else if (Convert.ToDateTime(dbAcDateFrom) > Convert.ToDateTime(varServEndDate))
                                                    varServEndDate = dbAcDateFrom;


                                                if (String.IsNullOrEmpty(dbAcDateTo) == true)
                                                    varServEndDate = cleanDate(varMonthEnd);
                                                else if (Convert.ToDateTime(varServEndDate) < Convert.ToDateTime(dbAcDateTo))
                                                    varServEndDate = dbAcDateTo;

                                                if (dbAccomType == "1")
                                                    varAccomST = true;
                                                if (dbAccomType == "2")
                                                    varAccomMT = true;
                                                if (dbAccomType == "3")
                                                    varAccomLT = true;
                                            }
                                        }

                                    } // Same support period
                                } // Accom Loop

                                //5d:
                                //Support Services
                                //C13.004.02 | Short term accom provided but not in accom table
                                //C13.004.03 | Medium term accom provided but not in accom table
                                //C13.004.04 | Long term accom provided but not in accom table

                                // Ensure service only processed once
                                varServNeedM = "**";
                                varServProvideM = "**";
                                varServArrangeM = "**";

                                // Loop through services records
                                foreach (var s in result7.Entities)
                                {
                                    // We need to get the entity id for the client field for comparisons
                                    if (s.Attributes.Contains("new_supportperiod"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)s.Attributes["new_supportperiod"];
                                        dbSvSupportPeriod = getEntity.Id.ToString();
                                    }
                                    else if (s.FormattedValues.Contains("new_supportperiod"))
                                        dbSvSupportPeriod = s.FormattedValues["new_supportperiod"];
                                    else
                                        dbSvSupportPeriod = "";

                                    // Need to see if same support period
                                    if (dbPalmClientSupportId == dbSvSupportPeriod)
                                    {
                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (s.FormattedValues.Contains("new_entrydate"))
                                            dbSvEntryDate = s.FormattedValues["new_entrydate"];
                                        else if (s.Attributes.Contains("new_entrydate"))
                                            dbSvEntryDate = s.Attributes["new_entrydate"].ToString();
                                        else
                                            dbSvEntryDate = "";

                                        // Convert date from American format to Australian format
                                        dbSvEntryDate = cleanDateAM(dbSvEntryDate);

                                        if (s.FormattedValues.Contains("new_description"))
                                            dbSvDescription = s.FormattedValues["new_description"];
                                        else if (s.Attributes.Contains("new_description"))
                                            dbSvDescription = s.Attributes["new_description"].ToString();
                                        else
                                            dbSvDescription = "";

                                        if (s.FormattedValues.Contains("new_servneed"))
                                            dbServNeed = s.FormattedValues["new_servneed"].ToString();
                                        else if (s.Attributes.Contains("new_servneed"))
                                            dbServNeed = s.Attributes["new_servneed"].ToString();
                                        else
                                            dbServNeed = "";

                                        // Wrap asterisks around values for better comparison
                                        dbServNeed = getMult(dbServNeed);

                                        if (s.FormattedValues.Contains("new_servprovide"))
                                            dbServProvide = s.FormattedValues["new_servprovide"];
                                        else if (s.Attributes.Contains("new_servprovide"))
                                            dbServProvide = s.Attributes["new_servprovide"].ToString();
                                        else
                                            dbServProvide = "";

                                        // Wrap asterisks around values for better comparison
                                        dbServProvide = getMult(dbServProvide);

                                        if (s.FormattedValues.Contains("new_servarrange"))
                                            dbServArrange = s.FormattedValues["new_servarrange"];
                                        else if (s.Attributes.Contains("new_servarrange"))
                                            dbServArrange = s.Attributes["new_servarrange"].ToString();
                                        else
                                            dbServArrange = "";

                                        // Wrap asterisks around values for better comparison
                                        dbServArrange = getMult(dbServArrange);

                                        // Extreme Weather Fix
                                        if (dbServNeed.IndexOf("*Extreme Weather Program*") > -1)
                                            dbServNeed = dbServNeed.Replace("Extreme Weather Program", "Health/medical services");
                                        if (dbServProvide.IndexOf("*Extreme Weather Program*") > -1)
                                            dbServProvide = dbServProvide.Replace("Extreme Weather Program", "Health/medical services");
                                        if (dbServArrange.IndexOf("*Extreme Weather Program*") > -1)
                                            dbServArrange = dbServArrange.Replace("Extreme Weather Program", "Health/medical services");

                                        // Family Violence Fix
                                        if (dbServNeed.IndexOf("*Assistance for domestic/family violence*") > -1)
                                            dbServNeed = dbServNeed.Replace("Assistance for domestic/family violence", "Assistance for family/domestic violence – victim support services");
                                        if (dbServProvide.IndexOf("*Assistance for domestic/family violence*") > -1)
                                            dbServProvide = dbServProvide.Replace("Assistance for domestic/family violence", "Assistance for family/domestic violence – victim support services");
                                        if (dbServArrange.IndexOf("*Assistance for domestic/family violence*") > -1)
                                            dbServArrange = dbServArrange.Replace("Assistance for domestic/family violence", "Assistance for family/domestic violence – victim support services");

                                        // Value needs to be 14 for earlier support periods
                                        if ((DateTime.Compare(Convert.ToDateTime(dbStartDate), Convert.ToDateTime("1-Jul-2019")) < 0) || varMonthStart < Convert.ToDateTime("1-Jul-2019"))
                                        {
                                            if (dbServNeed.IndexOf("*Assistance for family/domestic violence - victim support services*") > -1)
                                                dbServNeed = dbServNeed.Replace("Assistance for family/domestic violence - victim support services", "14");
                                            if (dbServNeed.IndexOf("*Assistance for family/domestic violence - perpetrator support service*") > -1)
                                                dbServNeed = dbServNeed.Replace("Assistance for family/domestic violence - perpetrator support service", "14");
                                            if (dbServProvide.IndexOf("*Assistance for family/domestic violence - victim support services*") > -1)
                                                dbServProvide = dbServProvide.Replace("Assistance for family/domestic violence - victim support services", "14");
                                            if (dbServProvide.IndexOf("*Assistance for family/domestic violence - perpetrator support service*") > -1)
                                                dbServProvide = dbServProvide.Replace("Assistance for family/domestic violence - perpetrator support service", "14");
                                            if (dbServArrange.IndexOf("*Assistance for family/domestic violence - victim support services*") > -1)
                                                dbServArrange = dbServArrange.Replace("Assistance for family/domestic violence - victim support services", "14");
                                            if (dbServArrange.IndexOf("*Assistance for family/domestic violence - perpetrator support service*") > -1)
                                                dbServArrange = dbServArrange.Replace("Assistance for family/domestic violence - perpetrator support service", "14");
                                        }

                                        // Process if valid service date
                                        if (String.IsNullOrEmpty(dbSvEntryDate) == false)
                                        {
                                            // Check if date falls within period
                                            if (Convert.ToDateTime(dbSvEntryDate) >= varMonthStart && Convert.ToDateTime(dbSvEntryDate) <= varMonthEnd)
                                            {
                                                //Get service start date and end date
                                                if (String.IsNullOrEmpty(varServStartDate) == true)
                                                    varServStartDate = dbSvEntryDate;

                                                if (String.IsNullOrEmpty(varServEndDate) == true)
                                                    varServEndDate = dbSvEntryDate;
                                                else if (Convert.ToDateTime(varServEndDate) < Convert.ToDateTime(dbSvEntryDate))
                                                    varServEndDate = dbSvEntryDate;

                                                // Loop through drop down list
                                                foreach (var d in result2.Entities)
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

                                                    

                                                    if (d.FormattedValues.Contains("new_shor"))
                                                        varSHOR = d.FormattedValues["new_shor"];
                                                    else if (d.Attributes.Contains("new_shor"))
                                                        varSHOR = d.Attributes["new_shor"].ToString();
                                                    else
                                                        varSHOR = "";



                                                    // Process if service
                                                    if (varType == "service")
                                                    {

                                                        // Ensure numeric
                                                        varSHOR = cleanString(varSHOR, "number");
                                                        if (String.IsNullOrEmpty(varSHOR) == true)
                                                            varSHOR = "0";

                                                        Int32.TryParse(varSHOR, out varCheckInt);

                                                        // If the service is needed, provided or arranged, add the service as needed
                                                        if (dbServNeed.IndexOf("*" + varDesc + "*") > -1 || dbServProvide.IndexOf("*" + varDesc + "*") > -1 || dbServArrange.IndexOf("*" + varDesc + "*") > -1)
                                                        {
                                                            // Dont add needed for accom as this is done below
                                                            //if ((varCheckInt > 3 || (varCheckInt == 1 && varAccomST == false) || (varCheckInt == 2 && varAccomMT == false) || (varCheckInt == 3 && varAccomLT == false)) && varServNeedM.IndexOf("*" + varDesc + "*") == -1)
                                                            //{
                                                                //Services Count
                                                                varSScount++;
                                                                varServNeedM += "*" + varDesc + "*"; // Append to string to stop being processed twice

                                                                varServDone = 1; // At least one service provided
                                                                sbSHORServices.AppendLine("            <SP_CP_Support_Services>");
                                                                sbSHORServices.AppendLine("               <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                                sbSHORServices.AppendLine("               <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                                sbSHORServices.AppendLine("               <Type_Of_Service_Activity>" + varCheckInt + "</Type_Of_Service_Activity>");
                                                                sbSHORServices.AppendLine("               <Service_Activity_Outcome>1</Service_Activity_Outcome>");
                                                                sbSHORServices.AppendLine("            </SP_CP_Support_Services>");
                                                           // }
                                                        }

                                                        // If the service is provided add the service (if not already)
                                                        // Dont add provided for accom as this is done below
                                                        if (dbServProvide.IndexOf("*" + varDesc + "*") > -1 && varCheckInt > 3 && varServProvideM.IndexOf("*" + varDesc + "*") == -1)
                                                        {
                                                            //Services Count
                                                            varSScount++;
                                                            varServProvideM += "*" + varDesc + "*"; // Append to string to stop being processed twice

                                                            varServDone = 1; // At least one service provided
                                                            sbSHORServices.AppendLine("            <SP_CP_Support_Services>");
                                                            sbSHORServices.AppendLine("               <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                            sbSHORServices.AppendLine("               <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                            sbSHORServices.AppendLine("               <Type_Of_Service_Activity>" + varCheckInt + "</Type_Of_Service_Activity>");
                                                            sbSHORServices.AppendLine("               <Service_Activity_Outcome>2</Service_Activity_Outcome>");
                                                            sbSHORServices.AppendLine("            </SP_CP_Support_Services>");
                                                        }

                                                        // If the service is arranged add the service (if not already)
                                                        if (dbServArrange.IndexOf("*" + varDesc + "*") > -1 && varServArrangeM.IndexOf("*" + varDesc + "*") == -1)
                                                        {
                                                            //Services Count
                                                            varSScount++;
                                                            varServArrangeM += "*" + varDesc + "*"; // Append to string to stop being processed twice

                                                            varServDone = 1; // At least one service provided
                                                            sbSHORServices.AppendLine("            <SP_CP_Support_Services>");
                                                            sbSHORServices.AppendLine("               <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                            sbSHORServices.AppendLine("               <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                            sbSHORServices.AppendLine("               <Type_Of_Service_Activity>" + varCheckInt + "</Type_Of_Service_Activity>");
                                                            sbSHORServices.AppendLine("               <Service_Activity_Outcome>3</Service_Activity_Outcome>");
                                                            sbSHORServices.AppendLine("            </SP_CP_Support_Services>");
                                                        }
                                                    } // Type is service

                                                } //Drop down Loop

                                                // 14 Fix for old extracts
                                                if ((dbServNeed.IndexOf("*14*") > -1 || dbServProvide.IndexOf("*14*") > -1 || dbServArrange.IndexOf("*14*") > -1) && varServNeedM.IndexOf("*14*") == -1)
                                                {
                                                    //Services Count
                                                    varSScount++;
                                                    varServNeedM += "*14*";

                                                    varServDone = 1;
                                                    sbSHORServices.AppendLine("            <SP_CP_Support_Services>");
                                                    sbSHORServices.AppendLine("               <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                    sbSHORServices.AppendLine("               <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                    sbSHORServices.AppendLine("               <Type_Of_Service_Activity>14</Type_Of_Service_Activity>");
                                                    sbSHORServices.AppendLine("               <Service_Activity_Outcome>1</Service_Activity_Outcome>");
                                                    sbSHORServices.AppendLine("            </SP_CP_Support_Services>");
                                                }

                                                if (dbServProvide.IndexOf("*14*") > -1 && varServProvideM.IndexOf("*14*") == -1)
                                                {
                                                    //Services Count
                                                    varSScount++;
                                                    varServProvideM += "*14*";

                                                    varServDone = 1;
                                                    sbSHORServices.AppendLine("            <SP_CP_Support_Services>");
                                                    sbSHORServices.AppendLine("               <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                    sbSHORServices.AppendLine("               <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                    sbSHORServices.AppendLine("               <Type_Of_Service_Activity>14</Type_Of_Service_Activity>");
                                                    sbSHORServices.AppendLine("               <Service_Activity_Outcome>2</Service_Activity_Outcome>");
                                                    sbSHORServices.AppendLine("            </SP_CP_Support_Services>");
                                                }

                                                if (dbServArrange.IndexOf("*14*") > -1 && varServArrangeM.IndexOf("*14*") == -1)
                                                {
                                                    //Services Count
                                                    varSScount++;
                                                    varServArrangeM += "*14*";

                                                    varServDone = 1;
                                                    sbSHORServices.AppendLine("            <SP_CP_Support_Services>");
                                                    sbSHORServices.AppendLine("               <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                    sbSHORServices.AppendLine("               <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                    sbSHORServices.AppendLine("               <Type_Of_Service_Activity>14</Type_Of_Service_Activity>");
                                                    sbSHORServices.AppendLine("               <Service_Activity_Outcome>3</Service_Activity_Outcome>");
                                                    sbSHORServices.AppendLine("            </SP_CP_Support_Services>");
                                                }
                                                //End of 14 Fix

                                            } //Date falls in range

                                        } // Date not Null

                                    } // Same support period
                                } // Services Loop

                                //Make sure service start date is in the current month, and is not before the support period start date
                                if (String.IsNullOrEmpty(varServStartDate) == false)
                                {
                                    if (Convert.ToDateTime(varServStartDate) <= Convert.ToDateTime(dbStartDate))
                                        varServStartDate = dbStartDate;

                                    if (Convert.ToDateTime(varServStartDate) <= varMonthStart)
                                        varServStartDate = cleanDate(varMonthStart);
                                }
                                else if (dbStOngoing != "3")
                                {
                                    // Alert of no services if ongoing is not 3, and service date is empty
                                    varServStartDate = "1-Jan-1990";
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Service Start Date");
                                }

                                //Make sure service end date is in the current month, and is not before the support period start date
                                if (String.IsNullOrEmpty(varServEndDate) == false)
                                {
                                    if (Convert.ToDateTime(varServEndDate) >= Convert.ToDateTime(varMonthEnd))
                                        varServEndDate = cleanDate(varMonthEnd);

                                    if (Convert.ToDateTime(varServEndDate) <= Convert.ToDateTime(dbStartDate))
                                        varServEndDate = dbStartDate;

                                    if (Convert.ToDateTime(varServEndDate) <= Convert.ToDateTime(varServStartDate))
                                        varServEndDate = varServStartDate;
                                }
                                else if (dbStOngoing != "3")
                                {
                                    // Alert of no services if ongoing is not 3, and service date is empty
                                    varServEndDate = "1-Jan-1990";
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Service End Date");
                                }

                                //C12.004.02 | Error if submission is original and first date of service is not start date
                                if (String.IsNullOrEmpty(dbStartDate) == false && String.IsNullOrEmpty(varServStartDate) == false)
                                {
                                    if (varSubmission == 0 && (dbStartDate != varServStartDate))
                                        varServStartDate = dbStartDate;

                                    if (Convert.ToDateTime(varServStartDate) > Convert.ToDateTime(varServEndDate))
                                        varServStartDate = varServEndDate;

                                    //Start date fix
                                    //if (Convert.ToDateTime(varServStartDate) < varMonthStart)
                                    //varError += "Invalid Data: (" + dbClientNumber + ") " + dbFirstName + " " + dbSurname + " - This is the first status update but the start date is in an earlier period. You will need to enter a SHOR start date in the support period page relevant for this period";
                                }

                                // Get ongoing value based on end date if not 3
                                if (dbStOngoing != "3")
                                {
                                    if (String.IsNullOrEmpty(dbEndDate) == false)
                                    {
                                        if (Convert.ToDateTime(dbEndDate) >= varMonthEnd)
                                            dbStOngoing = "1";
                                        else
                                            dbStOngoing = "2";
                                    }
                                    else
                                        dbStOngoing = "1";
                                }

                                //[SHOR](112) | Ongoing support period indicator
                                if (!Int32.TryParse(dbStOngoing, out varCheckInt))
                                    dbStOngoing = "1";

                                //[SHOR](113) | Labour fource status
                                if (!Int32.TryParse(dbStLabourForce, out varCheckInt))
                                    dbStLabourForce = "99";

                                //[SHOR](114) | FT PT status
                                if (dbStFtPt == "Full-time")
                                    dbStFtPt = "1";
                                else if (dbStFtPt == "Part-time")
                                    dbStFtPt = "2";
                                else if (dbStFtPt == "Don't know")
                                    dbStFtPt = "99";
                                else if (dbStFtPt == "Not applicable")
                                    dbStFtPt = "0";
                                else
                                    dbStFtPt = "99";

                                //[SHOR](115) | Student indicator
                                if (dbStStudInd == "Yes")
                                    dbStStudInd = "1";
                                else if (dbStStudInd == "No")
                                    dbStStudInd = "2";
                                else if (dbStStudInd == "Don't know")
                                    dbStStudInd = "99";
                                else
                                    dbStStudInd = "99";

                                //[SHOR](116) | Student type
                                if (!Int32.TryParse(dbStStudType, out varCheckInt))
                                    dbStStudType = "99";

                                //[SHOR](117) | Awaiting government payment
                                if (dbStAwaitGovt == "Yes")
                                    dbStAwaitGovt = "1";
                                else if (dbStAwaitGovt == "No")
                                    dbStAwaitGovt = "2";
                                else if (dbStAwaitGovt == "Don't know")
                                    dbStAwaitGovt = "99";
                                else if (dbStAwaitGovt == "Not applicable")
                                    dbStAwaitGovt = "0";
                                else
                                    dbStAwaitGovt = "99";

                                //[SHOR](118) | Living Arrangements
                                if (!Int32.TryParse(dbStLivingArrange, out varCheckInt))
                                    dbStLivingArrange = "99";

                                //[SHOR](119) | Occupancy
                                if (!Int32.TryParse(dbStOccupancy, out varCheckInt))
                                    dbStOccupancy = "99";

                                //[SHOR](120) | Residential
                                if (!Int32.TryParse(dbStResidential, out varCheckInt))
                                    dbStResidential = "99";

                                //[SHOR](121) | Tenure
                                if (!Int32.TryParse(dbStTenure, out varCheckInt))
                                    dbStTenure = "99";

                                //[SHOR](122) | Income
                                if (!Int32.TryParse(dbStIncome, out varCheckInt))
                                    dbStIncome = "99";

                                //[SHOR](123) | Child care order
                                if (!Int32.TryParse(dbStCareOrder, out varCheckInt))
                                    dbStCareOrder = "99";

                                //[SHOR](124) | Case management plan in place
                                if (dbStCaseMgt == "No")
                                    dbStCaseMgt = "2";
                                else
                                    dbStCaseMgt = "1";

                                //[SHOR](124) | Reason no case management plan in place
                                if (!Int32.TryParse(dbStCaseMgtReason, out varCheckInt))
                                    dbStCaseMgtReason = "0";

                                if (dbStCaseMgt == "2" && dbStCaseMgtReason == "0")
                                    dbStCaseMgtReason = "8";


                                //[SHOR](125) | Reason no case management plan in place other
                                if (String.IsNullOrEmpty(dbStCaseMgtReasonOth) == true)
                                    dbStCaseMgtReasonOth = "Reason not stated";
                                else if (dbStCaseMgtReasonOth.Length > 100)
                                    dbStCaseMgtReasonOth = "Reason not stated";
                                else
                                    dbStCaseMgtReasonOth = removeString(dbStCaseMgtReasonOth, "link");

                                
                                if (dbStCaseMgtGoal == "All")
                                    dbStCaseMgtGoal = "4";
                                else if (dbStCaseMgtGoal == "Up to half")
                                    dbStCaseMgtGoal = "2";
                                else if (dbStCaseMgtGoal == "Half")
                                    dbStCaseMgtGoal = "2";
                                else if (dbStCaseMgtGoal == "More than half")
                                    dbStCaseMgtGoal = "3";
                                else if (dbStCaseMgtGoal == "Not at all")
                                    dbStCaseMgtGoal = "1";
                                else if (dbStCaseMgtGoal == "No case management plan")
                                    dbStCaseMgtGoal = "88";
                                else
                                    dbStCaseMgtGoal = "1";
                                
                                // Ensure cessation numeric
                                if (!Int32.TryParse(dbCessation, out varCheckInt))
                                    dbCessation = "99";

                                //[SHOR](NEW) | ADF
                                if (dbADF == "Yes")
                                    dbADF = "1";
                                else if (dbADF == "No")
                                    dbADF = "2";
                                else if (dbADF == "Don't Know")
                                    dbADF = "99";
                                else if (dbADF == "Not applicable")
                                    dbADF = "0";
                                else
                                    dbADF = "99";

                                
                                /*
                                if (varAge < 18)
                                    dbADF = "0";

                                if(varAge >= 18 && dbADF == "0")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is " + varAge + " but has 'Not applicable' as ADF value");
                                */

                                // NDIS indicator
                                if (dbNDIS == "Yes")
                                    dbNDIS = "1";
                                else if (dbNDIS == "No")
                                    dbNDIS = "2";
                                else
                                    dbNDIS = "99";


                                //New income rules
                                //C07.036.01, C07.036.03, C07.037.01, C07.037.03
                                if (Convert.ToDateTime(dbStartDate) >= Convert.ToDateTime("1-Jul-2017"))
                                {
                                    if (dbIncomeWkBef == "7" || dbIncomeWkBef == "8" || dbIncomeWkBef == "9")
                                        dbIncomeWkBef = "18";
                                    if (dbIncomePres == "7" || dbIncomePres == "8" || dbIncomePres == "9")
                                        dbIncomePres = "18";
                                }

                                //New income rules
                                //C12.016.01, C12.016.03
                                if (varMonthStart >= Convert.ToDateTime("1-Jul-2017"))
                                {
                                    if (dbStIncome == "7" || dbStIncome == "8" || dbStIncome == "9")
                                        dbStIncome = "18";
                                }


                                //No status update
                                if (varDoSupport == 0)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - No Status update");


                                //SHOR Validation rules for client details

                                //C07.003.02 | Start date is less than assistance request date
                                if (String.IsNullOrEmpty(dbStartDate) == false && String.IsNullOrEmpty(dbAssessed) == false)
                                {
                                    if (Convert.ToDateTime(dbStartDate) < Convert.ToDateTime(dbAssessed))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Start date is less than assessed date");
                                }

                                //C07.008.02 | Assessed date cant be greater than DOB
                                if (String.IsNullOrEmpty(dbDob) == false && String.IsNullOrEmpty(dbAssessed) == false)
                                {
                                    if (Convert.ToDateTime(dbDob) > Convert.ToDateTime(dbAssessed))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - date of birth is greater than than assessed date");
                                }

                                //C07.010.02 | Cant have 115 year old client
                                if (Convert.ToInt32(dbYearArrival) < varCurrentDate.AddYears(-116).Year)
                                    dbYearArrival = "9999";

                                //C07.010.03 | Year of arrival can't be greater than start date
                                if (dbYearArrival != "9999" && Convert.ToInt32(dbYearArrival) > Convert.ToDateTime(dbStartDate).Year)
                                    dbYearArrival = "9999";

                                //C07.010.04 | Year of arrival can't be less than date of birth
                                if (dbYearArrival != "0" && Convert.ToInt32(dbYearArrival) < Convert.ToDateTime(dbDob).Year)
                                    dbYearArrival = "9999";

                                //C07.012.04 | Year of arrival must be 0000 if born in Australia. It can't be 0000 if not Australia or unknown
                                if (dbCountry == "1101" || dbCountry == "1199")
                                    dbYearArrival = "0000";

                                //C07.012.05 | Year of arrival can't be 0000 if not born in Australia or unknown
                                if (dbYearArrival == "0000" && (dbCountry != "9999" && dbCountry != "1101" && dbCountry != "1199"))
                                    dbYearArrival = "9999";

                                //Year Arrival
                                //9999	Don’t know
                                //0000	Not applicable

                                //Country
                                //9999	Not applicable
                                //0000	Don’t know


                                //55 day fix
                                //W07.014.01 | Assistance request date can't be more than 55 days before start date
                                if (String.IsNullOrEmpty(dbStartDate) == false && String.IsNullOrEmpty(dbAssessed) == false)
                                {
                                    if (Convert.ToDateTime(dbAssessed) < Convert.ToDateTime(dbStartDate).AddDays(-55)) { 
                                        dbAssessed = dbStartDate;
                                    }
                                }

                                //C07.016.02 | If relationship to PUH is self and unit count = 1 then IDs must be the same
                                if (dbPuhRship == "1" && varPUHcount == 1 && (dbPuhId != dbPalmClientSupportId))
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Presenting unit id is different to own id but relationship is self");

                                //C07.016.03 | If relationship to PUH not self and unit count > 1 then IDs cant be the same
                                if (dbPuhRship != "1" && varPUHcount > 1 && dbPuhId == dbPalmClientSupportId)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Presenting unit is not self but has same id as self");

                                //C07.017.02 | If IDs are the same and unit count = 1 then relationship needs to be self
                                if (dbPuhId == dbPalmClientSupportId && varPUHcount == 1)
                                {
                                    dbPuhRshipOther = "";
                                    dbPuhRship = "1";
                                }

                                //C07.017.03 | If IDs are not the same and unit count > 1 then relationship cant be self
                                if (dbPuhId != dbPalmClientSupportId && varPUHcount > 1 && dbPuhRship == "1")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Presenting unit is not self and unit count greater than 1 but relationship is self");

                                //C07.017.04 | Can't be grandparent if less than 18 (can be grand parent if 19 though)
                                if (dbPuhRship == "10" && varAge < 18)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Relationship is grandparents but person us under 18");

                                //C07.019.02 / W07.019.01 | Unit count cant be 0 or greater than 15
                                if (varPUHcount == 0)
                                {
                                    //PUH Fix
                                    dbPuhId = dbPalmClientSupportId;
                                    dbPuhRshipOther = "";
                                    dbPuhRship = "1";

                                    varPUHcount = 1;
                                    //sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Presenting unit head count is 0");
                                }
                                else if (varPUHcount > 15)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Presenting unit head count is greater than 15");

                                //PUHID Fix - add support period id
                                if (dbPuhRship == "1")
                                {
                                    dbPuhId = dbPalmClientSupportId;
                                    varPrintPuhId = varPrintSpId;
                                }
                                else
                                {
                                    dbPuhId = varInPuh; // If you see "none" as puhid, comment out this line

                                    if (string.IsNullOrEmpty(dbPuhIdOld) == false)
                                        varPrintPuhId = dbPuhIdOld;
                                    else
                                        varPrintPuhId = dbPuhId;
                                }

                                // Remove dashes from presenting unit id for extract
                                varPuhId = "";
                                if (String.IsNullOrEmpty(varPrintPuhId) == false)
                                    varPuhId = varPrintPuhId.Replace("-", "");


                                //C07.019.03 | Cant have unit count > 1 if IDs are the same (obsolete)
                                //If varPUHcount > 1 AND dbPuhId = dbClientNumber Then
                                //varError += "Invalid Data: (" + dbClientNumber + ") " + dbFirstName + " " + dbSurname + " - Presenting unit head count is greater than 1 but id is same";

                                //C07.019.04 | If relationship not self, can't have unit count of 1
                                if (dbPuhRship != "1" && varPUHcount == 1)
                                {
                                    //varError += "Invalid Data: (" + dbClientNumber + ") " + dbFirstName + " " + dbSurname + " - Presenting unit head count is 1 but relationship is not self";
                                    dbPuhRshipOther = "";
                                    dbPuhRship = "1";
                                }

                                //PUHID Fix - add support period id
                                if (dbPuhRship == "1")
                                    dbPuhId = dbPalmClientSupportId;
                                else
                                    dbPuhId = varInPuh;

                                //C07.025.02 | Error if last perm residence is greater than 5 years and client is less than 4. Womb counts as permanent address?
                                if (varAge < 4 && dbTimePerm == "6")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is less than 4 years old but has been out of permanent residence for more than 5 years");

                                //C07.029.02 | Error if client awaiting govt payment but is listed as having govt payment
                                if (dbAwaitGovtWkBef == "1" && Convert.ToInt32(dbIncomeWkBef) <= 13)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is awaiting government payment but has income source as government payment");

                                //C07.030.02 | Error if client awaiting govt payment but is listed as having govt payment
                                if (dbAwaitGovtPres == "1" && Convert.ToInt32(dbIncomePres) <= 13)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is awaiting government payment but has income source as government payment");

                                //C07.033.02 | Error if student is Yes but type is not applicable
                                if (dbStudIndWkBef == "1" && dbStudTypeWkBef == "0")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is studying but student type is listed as not applicable");

                                //C07.033.03 | Error if student is No but type different to not applicable
                                if (dbStudIndWkBef == "2" && dbStudTypeWkBef != "0")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is not studying but student type is not listed as not applicable");

                                //C07.033.04 | Error if student is unknown but type different to unknown or not applicable
                                if (dbStudIndWkBef == "99" && dbStudTypeWkBef != "0" && dbStudTypeWkBef != "99")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client studying is not known but student type is not listed as not applicable / don't know");

                                //C07.034.02 | Error if student is Yes but type is not applicable
                                if (dbStudIndPres == "1" && dbStudTypePres == "0")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is studying but student type is listed as not applicable");

                                //C07.034.03 | Error if student is No but type different to not applicable
                                if (dbStudIndPres == "2" && dbStudTypePres != "0")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is not studying but student type is not listed as not applicable");

                                //C07.034.04 | Error if student is unknown but type different to unknown or not applicable
                                if (dbStudIndPres == "99" && dbStudTypePres != "0" && dbStudTypePres != "99")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client studying is not known but student type is not listed as not applicable / don't know");

                                //W07.035.01 | Error if greater than 18 but education at present different to not applicable
                                if (varAge > 18)
                                    dbEducationPres = "0";

                                //C07.035.02 | Error if student is Yes and neither is chosen
                                if (dbStudIndPres == "1" && dbEducationPres == "6")
                                    dbEducationPres = "0";

                                //C07.035.03 | Error if student is No and education different to nothing or not applicable
                                if (dbStudIndPres == "2" && dbEducationPres != "0" && dbEducationPres != "6")
                                    dbEducationPres = "0";

                                //C07.035.04 | Error if student type is dont know and education is different to dont know or not applicable
                                if (dbStudIndPres == "99" && dbEducationPres != "0" && dbEducationPres != "99")
                                    dbEducationPres = "99";

                                //C07.036.20 | Error if age less than 18 and income type = 1, 5, 8, 9 and start date < 1 July 2017
                                if (varAge < 18 && (dbIncomeWkBef == "1" || dbIncomeWkBef == "5" || dbIncomeWkBef == "8" || dbIncomeWkBef == "9") && Convert.ToDateTime(dbStartDate) < Convert.ToDateTime("1-Jul-2017"))
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client aged under 18 has invalid income type week before");

                                //C07.036.21 | Error if age less than 18 and income type = 1, 5 and start date >= 1 July 2017
                                if (varAge < 18 && (dbIncomeWkBef == "1" || dbIncomeWkBef == "5") && Convert.ToDateTime(dbStartDate) >= Convert.ToDateTime("1-Jul-2017"))
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client aged under 18 has invalid income type week before");

                                //W07.036.01 | Error if age less than 15 and income type <> 17
                                if (varAge < 15)
                                    dbIncomeWkBef = "17";

                                //W07.037.01 | Error if age less than 15 and income type <> 17
                                if (varAge < 15)
                                    dbIncomePres = "17";

                                //C07.037.20 | Error if age less than 18 and income type = 1, 5, 8, 9 and start date < 1 July 2017
                                if (varAge < 18 && (dbIncomePres == "1" || dbIncomePres == "5" || dbIncomePres == "8" || dbIncomePres == "9") && Convert.ToDateTime(dbStartDate) < Convert.ToDateTime("1-Jul-2017"))
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client aged under 18 has invalid income type week before");

                                //C07.037.21 | Error if age less than 18 and income type = 1, 5 and start date >= 1 July 2017
                                if (varAge < 18 && (dbIncomePres == "1" || dbIncomePres == "5") && Convert.ToDateTime(dbStartDate) >= Convert.ToDateTime("1-Jul-2017"))
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client aged under 18 has invalid income type week before");

                                //C07.038.20 | Error if age less than 15 and living dbangements is couple without children
                                if (varAge < 15 && dbLivingArrangeWkBef == "4")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " - Client aged under 15 is living as couple without children");

                                //C07.039.20 | Error if age less than 15 and living dbangements is couple without children
                                if (varAge < 15 && dbLivingArrangePres == "4")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client aged under 15 is living as couple without children");

                                //W07.040.01 | Error if age less than 15 and labour force status different to not applicable
                                if (varAge < 15)
                                    dbLabourForceWkBef = "0";

                                //W07.041.01 | Error if age less than 15 and labour force status different to not applicable
                                if (varAge < 15)
                                    dbLabourForcePres = "0";

                                //C07.042.02 | Error if employed and FT/PT status not applicable
                                if (dbLabourForceWkBef == "1" && dbFtPtWkBef == "0")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is employed but labour force status is not applicable");

                                //C07.042.03 | Error if unemployed or not in labour force and FT/PT status different to not applicable
                                if ((dbLabourForceWkBef == "2" || dbLabourForceWkBef == "3") && dbFtPtWkBef != "0")
                                    dbFtPtWkBef = "0";

                                //C07.042.04 | Error if labour force status not known and FT/PT status is known
                                if (dbLabourForceWkBef == "99" && (dbFtPtWkBef == "1" || dbFtPtWkBef == "2"))
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client labour force status is dont know but pt/ft status has a value");

                                //W07.042.01 | Error if labour force status not applicable and FT/PT status different to not applicable
                                if (dbLabourForceWkBef == "0")
                                    dbFtPtWkBef = "0";

                                //C07.043.02 | Error if employed and FT/PT status not applicable
                                if (dbLabourForcePres == "1" && dbFtPtPres == "0")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is employed but labour force status is not applicable");

                                //C07.043.03 | Error if unemployed or not in labour force and FT/PT status different to not applicable
                                if ((dbLabourForcePres == "2" || dbLabourForcePres == "3") && dbFtPtPres != "0")
                                    dbFtPtPres = "0";

                                //C07.043.04 | Error if labour force status not known and FT/PT status is known
                                if (dbLabourForcePres == "99" && (dbFtPtPres == "1" || dbFtPtPres == "2"))
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client labour force status is dont know but pt/ft status has a value");

                                //W07.043.01 | Error if labour force status not applicable and FT/PT status different to not applicable
                                if (dbLabourForcePres == "0")
                                    dbFtPtPres = "0";

                                varDoErr = false;

                                //C07.044.02
                                if ((dbOccupancyWkBef == "1" || dbOccupancyWkBef == "2") && (dbResidentialWkBef == "6" || dbResidentialWkBef == "7" || dbResidentialWkBef == "8" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "20"))
                                    varDoErr = true;

                                //C07.044.03
                                if ((dbOccupancyWkBef == "3") && (dbResidentialWkBef == "7" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "14" || dbResidentialWkBef == "15" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "19" || dbResidentialWkBef == "20"))
                                    varDoErr = true;

                                //C07.044.04
                                if ((dbOccupancyWkBef == "4") && (dbResidentialWkBef == "6" || dbResidentialWkBef == "7" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "20"))
                                    varDoErr = true;

                                //C07.044.05
                                if ((dbOccupancyWkBef == "5") && (dbResidentialWkBef == "6" || dbResidentialWkBef == "7" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "19" || dbResidentialWkBef == "20"))
                                    varDoErr = true;

                                //C07.044.30
                                if (dbOccupancyWkBef == "3" && dbResidentialWkBef != "1")
                                    varDoErr = true;


                                if (varDoErr == true)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " " + dbStartDate + " - Inconsistant values for occupancy and residential status");


                                varDoErr = false;

                                //C07.045.02
                                if ((dbOccupancyPres == "1" || dbOccupancyPres == "2") && (dbResidentialPres == "6" || dbResidentialPres == "7" || dbResidentialPres == "8" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "20"))
                                    varDoErr = true;

                                //C07.045.03
                                if ((dbOccupancyPres == "3") && (dbResidentialPres == "7" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "14" || dbResidentialPres == "15" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "19" || dbResidentialPres == "20"))
                                    varDoErr = true;

                                //C07.045.04
                                if ((dbOccupancyPres == "4") && (dbResidentialPres == "6" || dbResidentialPres == "7" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "20"))
                                    varDoErr = true;

                                //C07.045.05
                                if ((dbOccupancyPres == "5") && (dbResidentialPres == "6" || dbResidentialPres == "7" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "19" || dbResidentialPres == "20"))
                                    varDoErr = true;

                                //C07.045.30
                                if (dbOccupancyPres == "3" && dbResidentialPres != "1")
                                    varDoErr = true;


                                if (varDoErr == true)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Inconsistant values for occupancy and residential status");


                                varDoErr = false;

                                //C07.046.02
                                if ((dbResidentialWkBef == "6" || dbResidentialWkBef == "7" || dbResidentialWkBef == "8" || dbResidentialWkBef == "9" || dbResidentialWkBef == "10" || dbResidentialWkBef == "11" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "14" || dbResidentialWkBef == "15" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "18" || dbResidentialWkBef == "19" || dbResidentialWkBef == "20") && (dbTenureWkBef == "1" || dbTenureWkBef == "9"))
                                    varDoErr = true;

                                //C07.046.03
                                if ((dbResidentialWkBef == "6" || dbResidentialWkBef == "7" || dbResidentialWkBef == "8" || dbResidentialWkBef == "9" || dbResidentialWkBef == "11" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "18" || dbResidentialWkBef == "19" || dbResidentialWkBef == "20") && (dbTenureWkBef == "2" || dbTenureWkBef == "3" || dbTenureWkBef == "10" || dbTenureWkBef == "11"))
                                    varDoErr = true;

                                //C07.046.04
                                if ((dbResidentialWkBef == "6" || dbResidentialWkBef == "7" || dbResidentialWkBef == "8" || dbResidentialWkBef == "9" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "18" || dbResidentialWkBef == "19" || dbResidentialWkBef == "20") && (dbTenureWkBef == "4" || dbTenureWkBef == "12"))
                                    varDoErr = true;

                                //C07.046.05
                                if ((dbResidentialWkBef == "9" || dbResidentialWkBef == "11" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "14" || dbResidentialWkBef == "15" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "18" || dbResidentialWkBef == "19" || dbResidentialWkBef == "20") && (dbTenureWkBef == "5" || dbTenureWkBef == "13"))
                                    varDoErr = true;

                                //C07.046.06
                                if ((dbResidentialWkBef == "3" || dbResidentialWkBef == "5" || dbResidentialWkBef == "6" || dbResidentialWkBef == "7" || dbResidentialWkBef == "8" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "19" || dbResidentialWkBef == "20") && (dbTenureWkBef == "6" || dbTenureWkBef == "14"))
                                    varDoErr = true;

                                //C07.046.07
                                if ((dbResidentialWkBef == "6" || dbResidentialWkBef == "7" || dbResidentialWkBef == "8" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "18" || dbResidentialWkBef == "19" || dbResidentialWkBef == "20") && (dbTenureWkBef == "7" || dbTenureWkBef == "15"))
                                    varDoErr = true;

                                //C07.046.08
                                if ((dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "20") && (dbTenureWkBef == "8" || dbTenureWkBef == "16"))
                                    varDoErr = true;

                                //C07.046.09
                                if ((dbResidentialWkBef == "6" || dbResidentialWkBef == "7" || dbResidentialWkBef == "9" || dbResidentialWkBef == "10" || dbResidentialWkBef == "11" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "14" || dbResidentialWkBef == "15" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "18" || dbResidentialWkBef == "19" || dbResidentialWkBef == "20") && (dbTenureWkBef == "18" || dbTenureWkBef == "19" || dbTenureWkBef == "20"))
                                    varDoErr = true;

                                //C07.046.10
                                if ((dbResidentialWkBef == "7" || dbResidentialWkBef == "8" || dbResidentialWkBef == "10" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "14" || dbResidentialWkBef == "15" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "18" || dbResidentialWkBef == "19" || dbResidentialWkBef == "20") && (dbTenureWkBef == "17"))
                                    varDoErr = true;

                                //C07.046.11
                                if ((dbResidentialWkBef == "7" || dbResidentialWkBef == "8" || dbResidentialWkBef == "12" || dbResidentialWkBef == "13" || dbResidentialWkBef == "16" || dbResidentialWkBef == "17" || dbResidentialWkBef == "20") && (dbTenureWkBef == "21"))
                                    varDoErr = true;

                                //C07.046.20
                                //if ((dbResidentialWkBef == "3" || dbResidentialWkBef == "4" || dbResidentialWkBef == "5" || dbResidentialWkBef == "6" || dbResidentialWkBef == "7" || dbResidentialWkBef == "8") && (dbTenureWkBef != "22" && dbTenureWkBef != "99"))
                                //varDoErr = true;

                                if (varDoErr == true)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Inconsistant values for residential status and tenure");


                                varDoErr = false;

                                //C07.047.02
                                if ((dbResidentialPres == "6" || dbResidentialPres == "7" || dbResidentialPres == "8" || dbResidentialPres == "9" || dbResidentialPres == "10" || dbResidentialPres == "11" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "14" || dbResidentialPres == "15" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "18" || dbResidentialPres == "19" || dbResidentialPres == "20") && (dbTenurePres == "1" || dbTenurePres == "9"))
                                    varDoErr = true;

                                //C07.047.03
                                if ((dbResidentialPres == "6" || dbResidentialPres == "7" || dbResidentialPres == "8" || dbResidentialPres == "9" || dbResidentialPres == "11" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "18" || dbResidentialPres == "19" || dbResidentialPres == "20") && (dbTenurePres == "2" || dbTenurePres == "3" || dbTenurePres == "10" || dbTenurePres == "11"))
                                    varDoErr = true;

                                //C07.047.04
                                if ((dbResidentialPres == "6" || dbResidentialPres == "7" || dbResidentialPres == "8" || dbResidentialPres == "9" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "18" || dbResidentialPres == "19" || dbResidentialPres == "20") && (dbTenurePres == "4" || dbTenurePres == "12"))
                                    varDoErr = true;

                                //C07.047.05
                                if ((dbResidentialPres == "9" || dbResidentialPres == "11" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "14" || dbResidentialPres == "15" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "18" || dbResidentialPres == "19" || dbResidentialPres == "20") && (dbTenurePres == "5" || dbTenurePres == "13"))
                                    varDoErr = true;

                                //C07.047.06
                                if ((dbResidentialPres == "3" || dbResidentialPres == "5" || dbResidentialPres == "6" || dbResidentialPres == "7" || dbResidentialPres == "8" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "19" || dbResidentialPres == "20") && (dbTenurePres == "6" || dbTenurePres == "14"))
                                    varDoErr = true;

                                //C07.047.07
                                if ((dbResidentialPres == "6" || dbResidentialPres == "7" || dbResidentialPres == "8" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "18" || dbResidentialPres == "19" || dbResidentialPres == "20") && (dbTenurePres == "7" || dbTenurePres == "15"))
                                    varDoErr = true;

                                //C07.047.08
                                if ((dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "20") && (dbTenurePres == "8" || dbTenurePres == "16"))
                                    varDoErr = true;

                                //C07.047.09
                                if ((dbResidentialPres == "6" || dbResidentialPres == "7" || dbResidentialPres == "9" || dbResidentialPres == "10" || dbResidentialPres == "11" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "14" || dbResidentialPres == "15" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "18" || dbResidentialPres == "19" || dbResidentialPres == "20") && (dbTenurePres == "18" || dbTenurePres == "19" || dbTenurePres == "20"))
                                    varDoErr = true;

                                //C07.047.10
                                if ((dbResidentialPres == "7" || dbResidentialPres == "8" || dbResidentialPres == "10" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "14" || dbResidentialPres == "15" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "18" || dbResidentialPres == "19" || dbResidentialPres == "20") && (dbTenurePres == "17"))
                                    varDoErr = true;

                                //C07.047.11
                                if ((dbResidentialPres == "7" || dbResidentialPres == "8" || dbResidentialPres == "12" || dbResidentialPres == "13" || dbResidentialPres == "16" || dbResidentialPres == "17" || dbResidentialPres == "20") && (dbTenurePres == "21"))
                                    varDoErr = true;

                                //C07.047.20
                                //if ((dbResidentialPres == "3" || dbResidentialPres == "4" || dbResidentialPres == "5" || dbResidentialPres == "6" || dbResidentialPres == "7" || dbResidentialPres == "8") && (dbTenurePres != "22" && dbTenurePres != "99"))
                                //varDoErr = true;

                                if (varDoErr == true)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Inconsistant values for residential status and tenure");


                                varDoErr = false;

                                //C07.048.02
                                if ((dbTenureWkBef == "19" || dbTenureWkBef == "20" || dbTenureWkBef == "22") && (dbOccupancyWkBef == "1" || dbOccupancyWkBef == "2"))
                                    varDoErr = true;

                                //C07.048.03
                                if ((dbTenureWkBef == "17" || dbTenureWkBef == "18" || dbTenureWkBef == "19" || dbTenureWkBef == "20") && (dbOccupancyWkBef == "3" || dbOccupancyWkBef == "4"))
                                    varDoErr = true;

                                //C07.048.04
                                if ((dbTenureWkBef == "1" || dbTenureWkBef == "2" || dbTenureWkBef == "3" || dbTenureWkBef == "4" || dbTenureWkBef == "5" || dbTenureWkBef == "6" || dbTenureWkBef == "7" || dbTenureWkBef == "8" || dbTenureWkBef == "15" || dbTenureWkBef == "17" || dbTenureWkBef == "18" || dbTenureWkBef == "19" || dbTenureWkBef == "20") && (dbOccupancyWkBef == "5"))
                                    varDoErr = true;

                                if (varDoErr == true)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Inconsistant values for tenure and occupancy");

                                varDoErr = false;

                                //C07.049.02
                                if ((dbTenurePres == "19" || dbTenurePres == "20" || dbTenurePres == "22") && (dbOccupancyPres == "1" || dbOccupancyPres == "2"))
                                    varDoErr = true;

                                //C07.049.03
                                if ((dbTenurePres == "17" || dbTenurePres == "18" || dbTenurePres == "19" || dbTenurePres == "20") && (dbOccupancyPres == "3" || dbOccupancyPres == "4"))
                                    varDoErr = true;

                                //C07.049.04
                                if ((dbTenurePres == "1" || dbTenurePres == "2" || dbTenurePres == "3" || dbTenurePres == "4" || dbTenurePres == "5" || dbTenurePres == "6" || dbTenurePres == "7" || dbTenurePres == "8" || dbTenurePres == "15" || dbTenurePres == "17" || dbTenurePres == "18" || dbTenurePres == "19" || dbTenurePres == "20") && (dbOccupancyPres == "5"))
                                    varDoErr = true;

                                if (varDoErr == true)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Inconsistant values for tenure and occupancy");

                                //W07.050.01 | Error if age greater than 18 and child care order different to not applicable
                                if (varAge >= 18)
                                    dbCareWkBef = "0";

                                //W07.051.01 | Error if age greater than 18 and child care order different to not applicable
                                if (varAge >= 18)
                                    dbCarePres = "0";

                                //W12.007.01 | Error if age less than 15 and status update labour force status different to not applicable
                                if (varAge < 15)
                                    dbStLabourForce = "0";

                                //C12.008.02 | Error if employed and status update FT/PT status not applicable
                                if (dbStLabourForce == "1" && dbStFtPt == "0")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is employed but labour force status is not applicable");

                                //C12.008.03 | Error if unemployed or not in labour force and status update FT/PT status different to not applicable
                                if ((dbStLabourForce == "2" || dbStLabourForce == "3") && dbStFtPt != "0")
                                    dbStFtPt = "0";

                                //C12.008.04 | Error if labour force status not known and status update FT/PT status is known
                                if (dbStLabourForce == "99" && (dbStFtPt == "1" || dbStFtPt == "2"))
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client labour force status is dont know but pt/ft status has a value");

                                //W12.008.01 | Error if labour force status not applicable and status update FT/PT status different to not applicable
                                if (dbStLabourForce == "0")
                                    dbStFtPt = "0";

                                //C12.010.02 | Error if student is Yes but type is not applicable
                                if (dbStStudInd == "1" && dbStStudType == "0")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is studying but student type is listed as not applicable");

                                //C12.010.03 | Error if student is No but type different to not applicable
                                if (dbStStudInd == "2" && dbStStudType != "0")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is not studying but status student type is not listed as not applicable");

                                //C12.010.04 | Error if student is unknown but type different to unknown or not applicable
                                if (dbStStudInd == "99" && dbStStudType != "0" && dbStStudType != "99")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client studying is not known but student type is not listed as not applicable / don't know");

                                //C12.011.02 | Error if client awaiting govt payment but is listed as having govt payment
                                if (dbStAwaitGovt == "1" && Convert.ToInt32(dbStIncome) <= 13)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client is awaiting government payment but has income source as government payment");

                                //C12.012.20
                                if (varAge < 15 && dbStLivingArrange == "4")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client aged under 15 is living as couple without children");

                                varDoErr = false;

                                //C12.014.02
                                if ((dbStOccupancy == "1" || dbStOccupancy == "2") && (dbStResidential == "6" || dbStResidential == "7" || dbStResidential == "8" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "20"))
                                    varDoErr = true;

                                //C12.014.03
                                if ((dbStOccupancy == "3") && (dbStResidential == "7" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "14" || dbStResidential == "15" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "19" || dbStResidential == "20"))
                                    varDoErr = true;

                                //C12.014.04
                                if ((dbStOccupancy == "4") && (dbStResidential == "6" || dbStResidential == "7" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "20"))
                                    varDoErr = true;

                                //C12.014.05
                                if ((dbStOccupancy == "5") && (dbStResidential == "6" || dbStResidential == "7" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "19" || dbStResidential == "20"))
                                    varDoErr = true;

                                //C12.014.30
                                if (dbStOccupancy == "3" && dbStResidential != "1")
                                    varDoErr = true;


                                if (varDoErr == true)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Inconsistant values for occupancy and residential status in Status Update");


                                varDoErr = false;

                                //C12.015.02
                                if ((dbStResidential == "6" || dbStResidential == "7" || dbStResidential == "8" || dbStResidential == "9" || dbStResidential == "10" || dbStResidential == "11" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "14" || dbStResidential == "15" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "18" || dbStResidential == "19" || dbStResidential == "20") && (dbStTenure == "1" || dbStTenure == "9"))
                                    varDoErr = true;

                                //C12.015.03
                                if ((dbStResidential == "6" || dbStResidential == "7" || dbStResidential == "8" || dbStResidential == "9" || dbStResidential == "11" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "18" || dbStResidential == "19" || dbStResidential == "20") && (dbStTenure == "2" || dbStTenure == "3" || dbStTenure == "10" || dbStTenure == "11"))
                                    varDoErr = true;

                                //C12.015.04
                                if ((dbStResidential == "6" || dbStResidential == "7" || dbStResidential == "8" || dbStResidential == "9" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "18" || dbStResidential == "19" || dbStResidential == "20") && (dbStTenure == "4" || dbStTenure == "12"))
                                    varDoErr = true;

                                //C12.015.05
                                if ((dbStResidential == "9" || dbStResidential == "11" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "14" || dbStResidential == "15" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "18" || dbStResidential == "19" || dbStResidential == "20") && (dbStTenure == "5" || dbStTenure == "13"))
                                    varDoErr = true;

                                //C12.015.06
                                if ((dbStResidential == "3" || dbStResidential == "5" || dbStResidential == "6" || dbStResidential == "7" || dbStResidential == "8" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "19" || dbStResidential == "20") && (dbStTenure == "6" || dbStTenure == "14"))
                                    varDoErr = true;

                                //C12.015.07
                                if ((dbStResidential == "6" || dbStResidential == "7" || dbStResidential == "8" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "18" || dbStResidential == "19" || dbStResidential == "20") && (dbStTenure == "7" || dbStTenure == "15"))
                                    varDoErr = true;

                                //C12.015.08
                                if ((dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "20") && (dbStTenure == "8" || dbStTenure == "16"))
                                    varDoErr = true;

                                //C12.015.09
                                if ((dbStResidential == "6" || dbStResidential == "7" || dbStResidential == "9" || dbStResidential == "10" || dbStResidential == "11" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "14" || dbStResidential == "15" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "18" || dbStResidential == "19" || dbStResidential == "20") && (dbStTenure == "18" || dbStTenure == "19" || dbStTenure == "20"))
                                    varDoErr = true;

                                //C12.015.10
                                if ((dbStResidential == "7" || dbStResidential == "8" || dbStResidential == "10" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "14" || dbStResidential == "15" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "18" || dbStResidential == "19" || dbStResidential == "20") && (dbStTenure == "17"))
                                    varDoErr = true;

                                //C12.015.11
                                if ((dbStResidential == "7" || dbStResidential == "8" || dbStResidential == "12" || dbStResidential == "13" || dbStResidential == "16" || dbStResidential == "17" || dbStResidential == "20") && (dbStTenure == "21"))
                                    varDoErr = true;

                                //C12.015.20
                                //if ((dbStResidential == "3" || dbStResidential == "4" || dbStResidential == "5" || dbStResidential == "6" || dbStResidential == "7" || dbStResidential == "8") && (dbStTenure != "22" && dbStTenure != "99"))
                                //varDoErr = true;

                                if (varDoErr == true)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Inconsistant values for residential status and tenure in Status Update");


                                //C12.016.20 | Error if age less than 18 and income type = 1, 5, 8, 9 and start date < 1 July 2017
                                if (varAge < 18 && (dbStIncome == "1" || dbStIncome == "5" || dbStIncome == "8" || dbStIncome == "9") && varMonthStart < Convert.ToDateTime("1-Jul-2017"))
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client aged under 18 has invalid income type in status update");

                                //C12.016.21 | Error if age less than 18 and income type = 1, 5 and start date >= 1 July 2017
                                if (varAge < 18 && (dbStIncome == "1" || dbStIncome == "5") && varMonthStart >= Convert.ToDateTime("1-Jul-2017"))
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Client aged under 18 has invalid income type in status update");


                                varDoErr = false;

                                //C12.013.02
                                if ((dbStTenure == "19" || dbStTenure == "20" || dbStTenure == "22") && (dbStOccupancy == "1" || dbStOccupancy == "2"))
                                    varDoErr = true;

                                //C12.013.03
                                if ((dbStTenure == "17" || dbStTenure == "18" || dbStTenure == "19" || dbStTenure == "20") && (dbStOccupancy == "3" || dbStOccupancy == "4"))
                                    varDoErr = true;

                                //C12.013.04
                                if ((dbStTenure == "1" || dbStTenure == "2" || dbStTenure == "3" || dbStTenure == "4" || dbStTenure == "5" || dbStTenure == "6" || dbStTenure == "7" || dbStTenure == "8" || dbStTenure == "15" || dbStTenure == "17" || dbStTenure == "18" || dbStTenure == "19" || dbStTenure == "20") && (dbStOccupancy == "5"))
                                    varDoErr = true;

                                if (varDoErr == true)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Inconsistant values for tenure and occupancy in Status Update");


                                //W12.016.01 | Error if age less than 15 and income type <> 17
                                if (varAge < 15)
                                    dbStIncome = "17";

                                //W12.017.01 | Error if age greater than 18 and child care order different to not applicable
                                if (varAge > 18)
                                    dbStCareOrder = "0";


                                //If no consent, then change the value of the consent fields
                                if (dbConsent == "2")
                                {
                                    dbCareWkBef = "0";
                                    dbCarePres = "0";
                                    dbStCareOrder = "0";
                                    dbCountry = "9999";
                                    dbDiagnosedMH = "0";
                                    dbIndigenous = "0";
                                    dbMHServicesRecd = "0";
                                    dbMentalIllnessInfo = "0";
                                    dbFacilities = "*0*";
                                }

                                // Null values unless July 2019
                                if (Convert.ToDateTime(dbStartDate) < Convert.ToDateTime("1-Jul-2019") || varMonthStart < Convert.ToDateTime("1-Jul-2019"))
                                {
                                    dbLanguage = "";
                                    dbSpeakProficient = "";
                                    dbNDIS = "";
                                }

                                //Glitch fix - change of consent
                                if (dbCountry == "9999")
                                    dbYearArrival = "0000";

                                /*
                                //Fixes prior to July 2015
                                if (Convert.ToDateTime(dbStartDate) < Convert.ToDateTime("1-Jul-2015"))
                                {
                                    varSurname = null;
                                    varFirstName = null;
                                    dbGender = null;
                                    dbDob = null;
                                    dbDobEst = null;
                                    dbYearArrival = null;
                                    dbIndigenous = null;
                                    dbCountry = null;
                                    dbConsent = null;
                                    dbAssessed = null; //Possible date error
                                    dbNewClient = null;
                                    varPuhId = null;
                                    dbPuhRship = null;
                                    //varPUHcount = null; //Error int
                                    dbSourceRef = null;
                                    dbPrimReason = null;
                                    dbPresLocality = null;
                                    dbPresPostcode = null;
                                    dbPresState = null;
                                    dbTimePerm = null;
                                    dbWkBefLocality = null;
                                    dbWkBefPostcode = null;
                                    dbWkBefState = null;
                                    dbAwaitGovtPres = null;
                                    dbAwaitGovtWkBef = null;
                                    dbStudIndPres = null;
                                    dbStudIndWkBef = null;
                                   // db
                                }*/


                                //[SHOR](23) | Original Submission
                                //varSubmission
                                //0	Original support period submission (Doesn't exist in previous SHOR period) *New client, first SHOR
                                //1	Resubmission of initial support period information (Modified information)
                                //2	Collection period information only (Client has been submitted before)

                                //Needs to be 0,1,2
                                //Set to 0 if no previous support period
                                //Set to 0 if previous support period was closed
                                //Set to 1 or 2 if in immediate previous support period and ongoing

                                //3:
                                //Client support periods
                                //Multiple

                                //Support Period Count
                                varSPcount++;

                                // Add support period data to the extract
                                sbExtractList.AppendLine("      <SP_Support_Period>");
                                sbExtractList.AppendLine("         <Organisation_ID>" + varAgencyId + "</Organisation_ID>"); //[SHOR](20)
                                sbExtractList.AppendLine("         <Support_Period_ID>" + varSuppId + "</Support_Period_ID>"); //[SHOR](21)
                                sbExtractList.AppendLine("         <Episode_Start_Date>" + cleanDateS(Convert.ToDateTime(dbStartDate)) + "</Episode_Start_Date>"); //[SHOR](22)
                                sbExtractList.AppendLine("         <SP_Submission_Ind>" + varSubmission + "</SP_Submission_Ind>"); //[SHOR](23)

                                // This data is only added for new or resubmitted support periods
                               // if (varSubmission != 2)
                                //{
                                    sbExtractList.AppendLine("         <Letters_Of_Family_Name>" + varSurname + "</Letters_Of_Family_Name>"); //[SHOR](24)
                                    sbExtractList.AppendLine("         <Letters_Of_Given_Name>" + varFirstName + "</Letters_Of_Given_Name>"); //[SHOR](25)
                                    sbExtractList.AppendLine("         <Sex>" + dbGender + "</Sex>"); //[SHOR](26)
                                    sbExtractList.AppendLine("         <Date_Of_Birth>" + cleanDateSBirthDate(Convert.ToDateTime(dbDob), dbDobEst) + "</Date_Of_Birth>"); //[SHOR](27)
                                    sbExtractList.AppendLine("         <Date_Of_Birth_Accuracy_Ind>" + dbDobEst + "</Date_Of_Birth_Accuracy_Ind>"); //[SHOR](28)
                                    sbExtractList.AppendLine("         <Year_Of_Arrival_In_Aust>" + dbYearArrival + "</Year_Of_Arrival_In_Aust>"); //[SHOR](29)

                                    // Format changed after 1 July 2019 to move it lower
                                    if (varMonthStart < Convert.ToDateTime("1-Jul-2019"))
                                    {
                                        if (varMonthStart >= Convert.ToDateTime("1-Jul-2017"))
                                            sbExtractList.AppendLine("         <ADF_Ind>" + dbADF + "</ADF_Ind>"); //[SHOR](NEW)
                                    }

                                    sbExtractList.AppendLine("         <Indigenous_Status>" + dbIndigenous + "</Indigenous_Status>"); //[SHOR](30)

                                    if (Int32.TryParse(dbCountry, out varCheckInt))
                                        sbExtractList.AppendLine("         <Country_Of_Birth>" + dbCountry.PadLeft(4,'0') + "</Country_Of_Birth>"); //[SHOR](31)
                                    else
                                        sbExtractList.AppendLine("         <Country_Of_Birth xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");

                                    sbExtractList.AppendLine("         <Consent_Obtained_Ind>" + dbConsent + "</Consent_Obtained_Ind>"); //[SHOR](32)
                                    sbExtractList.AppendLine("         <Assistance_Request_Date>" + cleanDateS(Convert.ToDateTime(dbAssessed)) + "</Assistance_Request_Date>"); //[SHOR](33)
                                    sbExtractList.AppendLine("         <New_Client_Ind>" + dbNewClient + "</New_Client_Ind>"); //[SHOR](34)
                                    sbExtractList.AppendLine("         <PULK_Support_Period_ID>" + varPuhId + "</PULK_Support_Period_ID>"); //[SHOR](35)
                                    sbExtractList.AppendLine("         <Relationship_To_PUH>" + dbPuhRship + "</Relationship_To_PUH>"); //[SHOR](36)

                                    if (String.IsNullOrEmpty(dbPuhRshipOther) == false)
                                        sbExtractList.AppendLine("         <Relationship_To_PUH_Other>" + dbPuhRshipOther + "</Relationship_To_PUH_Other>"); //[SHOR](37)
                                    else
                                        sbExtractList.AppendLine("         <Relationship_To_PUH_Other xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />"); //[SHOR](37)

                                    sbExtractList.AppendLine("         <Count_In_Presenting_Unit>" + varPUHcount + "</Count_In_Presenting_Unit>"); //[SHOR](38)
                                    sbExtractList.AppendLine("         <Formal_Referral_Source>" + dbSourceRef + "</Formal_Referral_Source>"); //[SHOR](39)
                                    sbExtractList.AppendLine("         <Assist_Main_Reason>" + dbPrimReason + "</Assist_Main_Reason>"); //[SHOR](40)

                                    if (String.IsNullOrEmpty(dbPresLocality) == false)
                                        sbExtractList.AppendLine("         <Locality_Most_Recent>" + dbPresLocality + "</Locality_Most_Recent>"); //[SHOR](41)
                                    else
                                        sbExtractList.AppendLine("         <Locality_Most_Recent xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />"); //[SHOR](41)

                                    if (String.IsNullOrEmpty(dbPresPostcode) == false)
                                        sbExtractList.AppendLine("         <Postcode_Most_Recent>" + dbPresPostcode + "</Postcode_Most_Recent>"); //[SHOR](42)
                                    else
                                        sbExtractList.AppendLine("         <Postcode_Most_Recent xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />"); //[SHOR](42)

                                    sbExtractList.AppendLine("         <State_Most_Recent>" + dbPresState + "</State_Most_Recent>"); //[SHOR](43)
                                    sbExtractList.AppendLine("         <Time_Since_Most_Recent_Addr>" + dbTimePerm + "</Time_Since_Most_Recent_Addr>"); //[SHOR](44)

                                    if (String.IsNullOrEmpty(dbWkBefLocality) == false)
                                        sbExtractList.AppendLine("         <Locality_Wkbef>" + dbWkBefLocality + "</Locality_Wkbef>"); //[SHOR](45)
                                    else
                                        sbExtractList.AppendLine("         <Locality_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />"); //[SHOR](45)

                                    if (String.IsNullOrEmpty(dbWkBefPostcode) == false)
                                        sbExtractList.AppendLine("         <Postcode_Wkbef>" + dbWkBefPostcode + "</Postcode_Wkbef>"); //[SHOR](46)
                                    else
                                        sbExtractList.AppendLine("         <Postcode_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />"); //[SHOR](46)

                                    sbExtractList.AppendLine("         <State_Wkbef>" + dbWkBefState + "</State_Wkbef>"); //[SHOR](47)

                                    // Format changed after 1 July 2019
                                    if (varMonthStart >= Convert.ToDateTime("1-Jul-2019"))
                                    {
                                        sbExtractList.AppendLine("         <ADF_Ind>" + dbADF + "</ADF_Ind>"); //[SHOR](NEW)
                                    }
                                    sbExtractList.AppendLine("         <Awaiting_Govt_Pymt_Ind_Wkbef>" + dbAwaitGovtWkBef + "</Awaiting_Govt_Pymt_Ind_Wkbef>"); //[SHOR](48)
                                    sbExtractList.AppendLine("         <Awaiting_Govt_Pymt_Ind_Present>" + dbAwaitGovtPres + "</Awaiting_Govt_Pymt_Ind_Present>"); //[SHOR](49)
                                    sbExtractList.AppendLine("         <Student_Ind_Wkbef>" + dbStudIndWkBef + "</Student_Ind_Wkbef>"); //[SHOR](50)
                                    sbExtractList.AppendLine("         <Student_Ind_Present>" + dbStudIndPres + "</Student_Ind_Present>"); //[SHOR](51)
                                    sbExtractList.AppendLine("         <Student_Type_Wkbef>" + dbStudTypeWkBef + "</Student_Type_Wkbef>"); //[SHOR](52)
                                    sbExtractList.AppendLine("         <Student_Type_Present>" + dbStudTypePres + "</Student_Type_Present>"); //[SHOR](53)
                                    sbExtractList.AppendLine("         <Education_At_Present>" + dbEducationPres + "</Education_At_Present>"); //[SHOR](54)
                                    sbExtractList.AppendLine("         <Source_Of_Income_Wkbef>" + dbIncomeWkBef + "</Source_Of_Income_Wkbef>"); //[SHOR](55)
                                    sbExtractList.AppendLine("         <Source_Of_Income_Present>" + dbIncomePres + "</Source_Of_Income_Present>"); //[SHOR](56)
                                    sbExtractList.AppendLine("         <Living_Arngmnt_Wkbef>" + dbLivingArrangeWkBef + "</Living_Arngmnt_Wkbef>"); //[SHOR](57)
                                    sbExtractList.AppendLine("         <Living_Arngmnt_Present>" + dbLivingArrangePres + "</Living_Arngmnt_Present>"); //[SHOR](58)
                                    sbExtractList.AppendLine("         <Labour_Force_Status_Wkbef>" + dbLabourForceWkBef + "</Labour_Force_Status_Wkbef>"); //[SHOR](59)
                                    sbExtractList.AppendLine("         <Labour_Force_Status_Present>" + dbLabourForcePres + "</Labour_Force_Status_Present>"); //[SHOR](60)
                                    sbExtractList.AppendLine("         <FT_PT_Status_Wkbef>" + dbFtPtWkBef + "</FT_PT_Status_Wkbef>"); //[SHOR](61)
                                    sbExtractList.AppendLine("         <FT_PT_Status_Present>" + dbFtPtPres + "</FT_PT_Status_Present>"); //[SHOR](62)
                                    sbExtractList.AppendLine("         <Residential_Wkbef>" + dbResidentialWkBef + "</Residential_Wkbef>"); //[SHOR](63)
                                    sbExtractList.AppendLine("         <Residential_Present>" + dbResidentialPres + "</Residential_Present>"); //[SHOR](64)
                                    sbExtractList.AppendLine("         <Tenure_Wkbef>" + dbTenureWkBef + "</Tenure_Wkbef>"); //[SHOR](65)
                                    sbExtractList.AppendLine("         <Tenure_Present>" + dbTenurePres + "</Tenure_Present>"); //[SHOR](66)
                                    sbExtractList.AppendLine("         <Occupancy_Wkbef>" + dbOccupancyWkBef + "</Occupancy_Wkbef>"); //[SHOR](67)
                                    sbExtractList.AppendLine("         <Occupancy_Present>" + dbOccupancyPres + "</Occupancy_Present>"); //[SHOR](68)
                                    sbExtractList.AppendLine("         <Care_And_Prot_Order_Wkbef>" + dbCareWkBef + "</Care_And_Prot_Order_Wkbef>"); //[SHOR](69)
                                    sbExtractList.AppendLine("         <Care_And_Prot_Order_Present>" + dbCarePres + "</Care_And_Prot_Order_Present>"); //[SHOR](70)
                                    sbExtractList.AppendLine("         <Diagnosed_Mental_Health>" + dbDiagnosedMH + "</Diagnosed_Mental_Health>"); //[SHOR](71)
                                    sbExtractList.AppendLine("         <Mental_Health_Services_Recd>" + dbMHServicesRecd + "</Mental_Health_Services_Recd>"); //[SHOR](72)
                                    sbExtractList.AppendLine("         <Mental_Illness_Info_Sources>" + dbMentalIllnessInfo + "</Mental_Illness_Info_Sources>"); //[SHOR](73)

                                    // New values from 1 July 2019
                                    if (varMonthStart >= Convert.ToDateTime("1-Jul-2019"))
                                    {
                                        if (String.IsNullOrEmpty(dbLanguage) == false)
                                            sbExtractList.AppendLine("         <Language_Spoken_At_Home>" + dbLanguage + "</Language_Spoken_At_Home>");
                                        else
                                            sbExtractList.AppendLine("         <Language_Spoken_At_Home xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");

                                        if (String.IsNullOrEmpty(dbSpeakProficient) == false)
                                            sbExtractList.AppendLine("         <Proficiency_In_Spoken_English>" + dbSpeakProficient + "</Proficiency_In_Spoken_English>");
                                        else
                                            sbExtractList.AppendLine("         <Proficiency_In_Spoken_English xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");

                                        if (String.IsNullOrEmpty(dbNDIS) == false)
                                            sbExtractList.AppendLine("         <NDIS_Ind>" + dbNDIS + "</NDIS_Ind>");
                                        else
                                            sbExtractList.AppendLine("         <NDIS_Ind xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    }

                               // } //varSubmission != 2
                               /* else
                                {
                                    // Null values for ongoing extract
                                    sbExtractList.AppendLine("         <Letters_Of_Family_Name xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Letters_Of_Given_Name xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Sex xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Date_Of_Birth xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Date_Of_Birth_Accuracy_Ind xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Year_Of_Arrival_In_Aust xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");

                                    // Format changed after 1 July 2019 to move it lower
                                    if (varMonthStart < Convert.ToDateTime("1-Jul-2019"))
                                    {
                                        if (varMonthStart >= Convert.ToDateTime("1-Jul-2017"))
                                            sbExtractList.AppendLine("         <ADF_Ind xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    }

                                    sbExtractList.AppendLine("         <Indigenous_Status xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Country_Of_Birth xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Consent_Obtained_Ind xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Assistance_Request_Date xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <New_Client_Ind xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <PULK_Support_Period_ID xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Relationship_To_PUH xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Relationship_To_PUH_Other xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Count_In_Presenting_Unit xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Formal_Referral_Source xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Assist_Main_Reason xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Locality_Most_Recent xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Postcode_Most_Recent xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <State_Most_Recent xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Time_Since_Most_Recent_Addr xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Locality_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Postcode_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <State_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");

                                    // Format changed after 1 July 2019
                                    if (varMonthStart >= Convert.ToDateTime("1-Jul-2019"))
                                        sbExtractList.AppendLine("         <ADF_Ind xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");

                                    sbExtractList.AppendLine("         <Awaiting_Govt_Pymt_Ind_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Awaiting_Govt_Pymt_Ind_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Student_Ind_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Student_Ind_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Student_Type_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Student_Type_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Education_At_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Source_Of_Income_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Source_Of_Income_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Living_Arngmnt_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Living_Arngmnt_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Labour_Force_Status_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Labour_Force_Status_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <FT_PT_Status_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <FT_PT_Status_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Residential_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Residential_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Tenure_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Tenure_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Occupancy_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Occupancy_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Care_And_Prot_Order_Wkbef xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Care_And_Prot_Order_Present xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Diagnosed_Mental_Health xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Mental_Health_Services_Recd xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("         <Mental_Illness_Info_Sources xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");

                                    // New values from 1 July 2019
                                    if (varMonthStart >= Convert.ToDateTime("1-Jul-2019"))
                                    {
                                        sbExtractList.AppendLine("         <Language_Spoken_At_Home xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                        sbExtractList.AppendLine("         <Proficiency_In_Spoken_English xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                        sbExtractList.AppendLine("         <NDIS_Ind xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    }

                                } //varSubmission = 2
                                */
                                // Add to client table
                                sbClientList.AppendLine("CLIENT," + varSuppId + "," + dbClient + "," + dbFirstName + "," + dbSurname + "," + dbStartDate + ",");

                                //4a:
                                //Collection period

                                //Collection period Count
                                varCPcount++;

                                // Create collection period part of the status update
                                sbExtractList.AppendLine("         <SP_CP_Collection_Period>");

                                // Do this part if status update ended last month
                                if (dbStOngoing == "3")
                                {

                                    sbExtractList.AppendLine("            <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                    sbExtractList.AppendLine("            <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");

                                    // Element no lnger required after July 2019
                                    if (varMonthStart < Convert.ToDateTime("1-Jul-2019"))
                                        sbExtractList.AppendLine("            <Episode_Start_Date>" + cleanDateS(Convert.ToDateTime(dbStartDate)) + "</Episode_Start_Date>");

                                    sbExtractList.AppendLine("            <First_Service_Contact_Date xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Last_Service_Provision_Date xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");

                                    sbExtractList.AppendLine("            <Ongoing_Support_Period_Ind>" + dbStOngoing + "</Ongoing_Support_Period_Ind>");

                                    sbExtractList.AppendLine("            <Labour_Force_Status_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <FT_PT_Status_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Student_Ind_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Student_Type_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Awaiting_Govt_Pymt_Ind_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Living_Arngmnt_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Occupancy_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Residential_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Tenure_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Source_Of_Income_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Care_And_Prot_Order_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Case_Mgmt_Plan_Ind_CP_End xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Reason_No_Case_Mgmt_Plan xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Reason_No_Case_Mgmt_Plan_Other xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Case_Mgmt_Plan_Goal_Status xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    sbExtractList.AppendLine("            <Service_Episode_End_Reason>" + dbCessation + "</Service_Episode_End_Reason>");

                                    //End 4a:
                                    sbExtractList.AppendLine("         </SP_CP_Collection_Period>");

                                }
                                else
                                {
                                    // Do this part if the status update did not end last month
                                    sbExtractList.AppendLine("            <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                    sbExtractList.AppendLine("            <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");

                                    // Element no longer required after July 2019
                                    if (varMonthStart < Convert.ToDateTime("1-Jul-2019"))
                                        sbExtractList.AppendLine("            <Episode_Start_Date>" + cleanDateS(Convert.ToDateTime(dbStartDate)) + "</Episode_Start_Date>");

                                    sbExtractList.AppendLine("            <First_Service_Contact_Date>" + cleanDateS(Convert.ToDateTime(varServStartDate)) + "</First_Service_Contact_Date>");
                                    sbExtractList.AppendLine("            <Last_Service_Provision_Date>" + cleanDateS(Convert.ToDateTime(varServEndDate)) + "</Last_Service_Provision_Date>");
                                    sbExtractList.AppendLine("            <Ongoing_Support_Period_Ind>" + dbStOngoing + "</Ongoing_Support_Period_Ind>");
                                    sbExtractList.AppendLine("            <Labour_Force_Status_CP_End>" + dbStLabourForce + "</Labour_Force_Status_CP_End>");
                                    sbExtractList.AppendLine("            <FT_PT_Status_CP_End>" + dbStFtPt + "</FT_PT_Status_CP_End>");
                                    sbExtractList.AppendLine("            <Student_Ind_CP_End>" + dbStStudInd + "</Student_Ind_CP_End>");
                                    sbExtractList.AppendLine("            <Student_Type_CP_End>" + dbStStudType + "</Student_Type_CP_End>");
                                    sbExtractList.AppendLine("            <Awaiting_Govt_Pymt_Ind_CP_End>" + dbStAwaitGovt + "</Awaiting_Govt_Pymt_Ind_CP_End>");
                                    sbExtractList.AppendLine("            <Living_Arngmnt_CP_End>" + dbStLivingArrange + "</Living_Arngmnt_CP_End>");
                                    sbExtractList.AppendLine("            <Occupancy_CP_End>" + dbStOccupancy + "</Occupancy_CP_End>");
                                    sbExtractList.AppendLine("            <Residential_CP_End>" + dbStResidential + "</Residential_CP_End>");
                                    sbExtractList.AppendLine("            <Tenure_CP_End>" + dbStTenure + "</Tenure_CP_End>");
                                    sbExtractList.AppendLine("            <Source_Of_Income_CP_End>" + dbStIncome + "</Source_Of_Income_CP_End>");
                                    sbExtractList.AppendLine("            <Care_And_Prot_Order_CP_End>" + dbStCareOrder + "</Care_And_Prot_Order_CP_End>");
                                    sbExtractList.AppendLine("            <Case_Mgmt_Plan_Ind_CP_End>" + dbStCaseMgt + "</Case_Mgmt_Plan_Ind_CP_End>");

                                    if (dbStCaseMgt == "2")
                                        sbExtractList.AppendLine("            <Reason_No_Case_Mgmt_Plan>" + dbStCaseMgtReason + "</Reason_No_Case_Mgmt_Plan>");
                                    else
                                        sbExtractList.AppendLine("            <Reason_No_Case_Mgmt_Plan>0</Reason_No_Case_Mgmt_Plan>");

                                    if (dbStCaseMgtReason == "8")
                                        sbExtractList.AppendLine("            <Reason_No_Case_Mgmt_Plan_Other>" + dbStCaseMgtReasonOth + "</Reason_No_Case_Mgmt_Plan_Other>");
                                    else
                                        sbExtractList.AppendLine("            <Reason_No_Case_Mgmt_Plan_Other xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");

                                    if (dbStCaseMgt == "2")
                                        sbExtractList.AppendLine("            <Case_Mgmt_Plan_Goal_Status>88</Case_Mgmt_Plan_Goal_Status>");
                                    else
                                        sbExtractList.AppendLine("            <Case_Mgmt_Plan_Goal_Status>" + dbStCaseMgtGoal + "</Case_Mgmt_Plan_Goal_Status>");

                                    if (dbStOngoing == "1")
                                        sbExtractList.AppendLine("            <Service_Episode_End_Reason xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");
                                    else
                                        sbExtractList.AppendLine("            <Service_Episode_End_Reason>" + dbCessation + "</Service_Episode_End_Reason>");


                                    //5a:
                                    //Accom Information
                                    varLatestDate = ""; // Reset data variable

                                    // Loop through accom list
                                    foreach (var s in result8.Entities)
                                    {
                                        // We need to get the entity id for the support period field for comparisons
                                        if (s.Attributes.Contains("new_supportperiod"))
                                        {
                                            // Get the entity id for the client using the entity reference object
                                            getEntity = (EntityReference)s.Attributes["new_supportperiod"];
                                            dbAcSupportPeriod = getEntity.Id.ToString();
                                        }
                                        else if (s.FormattedValues.Contains("new_supportperiod"))
                                            dbAcSupportPeriod = s.FormattedValues["new_supportperiod"];
                                        else
                                            dbAcSupportPeriod = "";

                                        // Need to see if same support period
                                        if (dbPalmClientSupportId == dbAcSupportPeriod)
                                        {
                                            // Process the data as follows:
                                            // If there is a formatted value for the field, use it
                                            // Otherwise if there is a literal value for the field, use it
                                            // Otherwise the value wasn't returned so set as nothing
                                            if (s.FormattedValues.Contains("new_datefrom"))
                                                dbAcDateFrom = s.FormattedValues["new_datefrom"];
                                            else if (s.Attributes.Contains("new_datefrom"))
                                                dbAcDateFrom = s.Attributes["new_datefrom"].ToString();
                                            else
                                                dbAcDateFrom = "";

                                            // Convert date from American format to Australian format
                                            dbAcDateFrom = cleanDateAM(dbAcDateFrom);

                                            if (s.FormattedValues.Contains("new_dateto"))
                                                dbAcDateTo = s.FormattedValues["new_dateto"];
                                            else if (s.Attributes.Contains("new_dateto"))
                                                dbAcDateTo = s.Attributes["new_dateto"].ToString();
                                            else
                                                dbAcDateTo = "";

                                            // Convert date from American format to Australian format
                                            dbAcDateTo = cleanDateAM(dbAcDateTo);

                                            if (String.IsNullOrEmpty(dbAcDateTo) == false && String.IsNullOrEmpty(dbAcDateFrom) == false)
                                            {
                                                // Add one day if dates the same
                                                if (Convert.ToDateTime(dbAcDateTo) == Convert.ToDateTime(dbAcDateFrom))
                                                    dbAcDateTo = cleanDateAM(Convert.ToDateTime(dbAcDateTo).AddDays(1).ToString());
                                            }

                                            if (s.FormattedValues.Contains("new_location"))
                                                dbAcLocation = s.FormattedValues["new_location"];
                                            else if (s.Attributes.Contains("new_location"))
                                                dbAcLocation = s.Attributes["new_location"].ToString();
                                            else
                                                dbAcLocation = "";

                                            if (s.FormattedValues.Contains("new_accomtype"))
                                                dbAccomType = s.FormattedValues["new_accomtype"];
                                            else if (s.Attributes.Contains("new_accomtype"))
                                                dbAccomType = s.Attributes["new_accomtype"].ToString();
                                            else
                                                dbAccomType = "";

                                            // Process if from date valid
                                            if (String.IsNullOrEmpty(dbAcDateFrom) == false)
                                            {
                                                // Get accom type
                                                if (dbAccomType == "Short term or emergency accommodation")
                                                    dbAccomType = "1";
                                                else if (dbAccomType == "Medium term/transitional accommodation")
                                                    dbAccomType = "2";
                                                else if (dbAccomType == "Long term accommodation")
                                                    dbAccomType = "3";
                                                else
                                                    dbAccomType = "";

                                                // Reset variable
                                                varValidDate = 0;
                                                // If both dates are valid and are active during the period then set flag
                                                if (String.IsNullOrEmpty(dbAcDateFrom) == false && String.IsNullOrEmpty(dbAcDateTo) == false)
                                                {
                                                    if (Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthEnd) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateTo) <= Convert.ToDateTime(varMonthEnd))
                                                        varValidDate = 1;

                                                    if (Convert.ToDateTime(dbAcDateFrom) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthEnd) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateTo) <= Convert.ToDateTime(varMonthEnd))
                                                        varValidDate = 1;

                                                    if (Convert.ToDateTime(dbAcDateFrom) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthEnd) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthEnd))
                                                        varValidDate = 1;

                                                    if (Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthEnd) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthStart) && Convert.ToDateTime(dbAcDateTo) >= Convert.ToDateTime(varMonthEnd))
                                                        varValidDate = 1;

                                                }
                                                // If only from date is valid and occurs during the period then set the flag
                                                else if (String.IsNullOrEmpty(dbAcDateFrom) == false)
                                                {
                                                    if (Convert.ToDateTime(dbAcDateFrom) <= Convert.ToDateTime(varMonthEnd))
                                                        varValidDate = 1;
                                                }

                                                //If accommodation exists and the date falls within this month, get the service start and end dates
                                                if (varValidDate == 1 && String.IsNullOrEmpty(dbAccomType) == false)
                                                {
                                                    varDoEnd = false; // Whether the accom ends during the period

                                                    // Get latest accom date if the accom has an end date
                                                    if (String.IsNullOrEmpty(dbAcDateTo) == false)
                                                    {
                                                        //C14.003.03 | Accom start date cant be earlier than previous end date
                                                        if (String.IsNullOrEmpty(varLatestDate) == false)
                                                        {
                                                            if (Convert.ToDateTime(dbAcDateTo) < Convert.ToDateTime(dbAcDateFrom))
                                                                sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Accom start date occurs before previous accom finishes");
                                                        }

                                                        //C14.004.03 | Accom end date cant be earlier than accom start date
                                                        if (Convert.ToDateTime(dbAcDateFrom) > Convert.ToDateTime(dbAcDateTo))
                                                            sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - Accom end date is less than accom start date");

                                                        //Set the Latest Date
                                                        varLatestDate = dbAcDateTo;
                                                    }
                                                    // If there is no end date then force accom to end at end of month
                                                    else
                                                    {
                                                        dbAcDateTo = cleanDate(varMonthEnd);
                                                        varLatestDate = dbAcDateTo;
                                                        varDoEnd = true;
                                                    }


                                                    //C14.003.02 | Accom start date cant be earlier than service start date
                                                    if (Convert.ToDateTime(dbAcDateFrom) < Convert.ToDateTime(varMonthStart))
                                                        dbAcDateFrom = cleanDate(varMonthStart);

                                                    //C14.003.02 | Accom end date cant be later than service end date
                                                    if (Convert.ToDateTime(dbAcDateTo) > Convert.ToDateTime(varMonthEnd))
                                                    {
                                                        dbAcDateTo = cleanDate(varMonthEnd);
                                                        varDoEnd = true;
                                                    }

                                                    //C14.003.02 | Accom start date cant be earlier than service start date
                                                    if (Convert.ToDateTime(dbAcDateFrom) < Convert.ToDateTime(dbStartDate))
                                                        dbAcDateFrom = cleanDate(Convert.ToDateTime(dbStartDate));

                                                    // Fix days minus glitch
                                                    if (Convert.ToDateTime(dbAcDateTo).AddDays(-1) >= varMonthStart)
                                                    {
                                                        //Insert accom into services
                                                        //Services Count
                                                        varSScount++;

                                                        varServDone = 1;
                                                        sbSHORAccomServ.AppendLine("            <SP_CP_Support_Services>");
                                                        sbSHORAccomServ.AppendLine("               <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                        sbSHORAccomServ.AppendLine("               <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                        sbSHORAccomServ.AppendLine("               <Type_Of_Service_Activity>" + dbAccomType + "</Type_Of_Service_Activity>");
                                                        sbSHORAccomServ.AppendLine("               <Service_Activity_Outcome>2</Service_Activity_Outcome>");
                                                        sbSHORAccomServ.AppendLine("            </SP_CP_Support_Services>");


                                                        varSScount++;
                                                        sbSHORAccomServ.AppendLine("            <SP_CP_Support_Services>");
                                                        sbSHORAccomServ.AppendLine("               <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                        sbSHORAccomServ.AppendLine("               <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                        sbSHORAccomServ.AppendLine("               <Type_Of_Service_Activity>" + dbAccomType + "</Type_Of_Service_Activity>");
                                                        sbSHORAccomServ.AppendLine("               <Service_Activity_Outcome>1</Service_Activity_Outcome>");
                                                        sbSHORAccomServ.AppendLine("            </SP_CP_Support_Services>");


                                                        //Accom Provided Count
                                                        varAPcount++;

                                                        sbExtractList.AppendLine("            <SP_CP_Accomm_Periods>");
                                                        sbExtractList.AppendLine("               <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                        sbExtractList.AppendLine("               <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                        sbExtractList.AppendLine("               <Accomm_Start_Date>" + cleanDateS(Convert.ToDateTime(dbAcDateFrom)) + "</Accomm_Start_Date>");

                                                        // Accom ends 1 day earlier, but if the accom is open it ends at the end of the month
                                                        if (varDoEnd == true)
                                                            sbExtractList.AppendLine("               <Accomm_End_Date>" + cleanDateS(Convert.ToDateTime(dbAcDateTo)) + "</Accomm_End_Date>");
                                                        else
                                                            sbExtractList.AppendLine("               <Accomm_End_Date>" + cleanDateS(Convert.ToDateTime(dbAcDateTo).AddDays(-1)) + "</Accomm_End_Date>");

                                                        sbExtractList.AppendLine("               <Accomm_Type>" + dbAccomType + "</Accomm_Type>");
                                                        sbExtractList.AppendLine("            </SP_CP_Accomm_Periods>");
                                                    }
                                                }
                                            }

                                        } // Same support period

                                    } // Accom Loop

                                    //5b:
                                    // Loop through financial data
                                    foreach (var s in result9.Entities)
                                    {
                                        // We need to get the entity id for the client field for comparisons
                                        if (s.Attributes.Contains("new_supportperiod"))
                                        {
                                            // Get the entity id for the client using the entity reference object
                                            getEntity = (EntityReference)s.Attributes["new_supportperiod"];
                                            dbFinSupportPeriod = getEntity.Id.ToString();
                                            varTest = dbFinSupportPeriod;
                                        }
                                        else if (s.FormattedValues.Contains("new_supportperiod"))
                                            dbFinSupportPeriod = s.FormattedValues["new_supportperiod"];
                                        else
                                            dbFinSupportPeriod = "";


                                        // Need to see if same support period
                                        if (dbPalmClientSupportId == dbFinSupportPeriod)
                                        {
                                            // Process the data as follows:
                                            // If there is a formatted value for the field, use it
                                            // Otherwise if there is a literal value for the field, use it
                                            // Otherwise the value wasn't returned so set as nothing
                                            if (s.FormattedValues.Contains("new_entrydate"))
                                                dbFinEntryDate = s.FormattedValues["new_entrydate"];
                                            else if (s.Attributes.Contains("new_entrydate"))
                                                dbFinEntryDate = s.Attributes["new_entrydate"].ToString();
                                            else
                                                dbFinEntryDate = "";

                                            // Convert date from American format to Australian format
                                            dbFinEntryDate = cleanDateAM(dbFinEntryDate);

                                            if (s.FormattedValues.Contains("new_amount"))
                                                dbFinAmount = s.FormattedValues["new_amount"];
                                            else if (s.Attributes.Contains("new_amount"))
                                                dbFinAmount = s.Attributes["new_amount"].ToString();
                                            else
                                                dbFinAmount = "";

                                            // Ensure amount numeric
                                            dbFinAmount = cleanString(dbFinAmount, "double");

                                            if (Double.TryParse(dbFinAmount, out varCheckDouble))
                                                dbFinAmount = ((int)varCheckDouble).ToString();
                                            else
                                                dbFinAmount = "0";

                                            if (s.FormattedValues.Contains("new_shor"))
                                                dbFinShor = s.FormattedValues["new_shor"];
                                            else if (s.Attributes.Contains("new_shor"))
                                                dbFinShor = s.Attributes["new_shor"].ToString();
                                            else
                                                dbFinShor = "";

                                            if (String.IsNullOrEmpty(dbFinShor) == true)
                                            {
                                                dbFinShor = "Other payment"; //No SHOR payment type selected from SHS financial type.
                                            }


                                            varSeeType = false;

                                            //Get the values from the drop down tables
                                            foreach (var d in result2.Entities)
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

                                                if (d.FormattedValues.Contains("new_shor"))
                                                    varSHOR = d.FormattedValues["new_shor"];
                                                else if (d.Attributes.Contains("new_shor"))
                                                    varSHOR = d.Attributes["new_shor"].ToString();
                                                else
                                                    varSHOR = "";

                                                varSHOR = cleanString(varSHOR, "number");
                                                if (String.IsNullOrEmpty(varSHOR) == true)
                                                    varSHOR = "0";

                                                // Get numeric version of financial type
                                                if (varType == "fintype" && varDesc == dbFinShor)
                                                {
                                                    dbFinShor = varSHOR;
                                                    varSeeType = true;
                                                    break;
                                                }

                                            } //k Loop

                                            // Reset if value not found
                                            if (varSeeType == false)
                                                dbFinShor = "";

                                            // Append to financial part of extract
                                            if (String.IsNullOrEmpty(dbFinShor) == false && dbFinAmount != "0")
                                            {
                                                //Financial Support Count
                                                varFScount++;

                                                sbExtractList.AppendLine("            <SP_CP_Financial_Support>");
                                                sbExtractList.AppendLine("               <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                sbExtractList.AppendLine("               <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                                sbExtractList.AppendLine("               <Financial_Assistance_Type>" + dbFinShor + "</Financial_Assistance_Type>");
                                                sbExtractList.AppendLine("               <Financial_Assistance_Amount>" + dbFinAmount + "</Financial_Assistance_Amount>");
                                                sbExtractList.AppendLine("            </SP_CP_Financial_Support>");

                                            }

                                        } // Same support period

                                    } // Financial Loop

                                    //5c:
                                    //Homeless History

                                    // Append homeless status update
                                    if (sbSHORHomeSt.Length > 0)
                                        sbExtractList.Append(sbSHORHomeSt.ToString());

                                    // Append services
                                    if (sbSHORServices.Length > 0)
                                        sbExtractList.Append(sbSHORServices.ToString());

                                    // Append Accom
                                    if (sbSHORAccomServ.Length > 0)
                                        sbExtractList.Append(sbSHORAccomServ.ToString());

                                    //C12.006.10 | Need to have a service per client
                                    if (varServDone == 0 && dbStOngoing != "3")
                                        sbErrorList.AppendLine("Invalid Data: (" + dbClient + ") " + dbFirstName + " " + dbSurname + " " + dbStartDate + " - There is no service against the client for this month");


                                    //End 4a:
                                    sbExtractList.AppendLine("         </SP_CP_Collection_Period>");

                                } // Ongoing != 3

                                // Only include for new or resubmits
                               // if (varSubmission != 2)
                               // {
                                    //4b:
                                    //Homeless_History_Mth | Only do if original submission or resubmission
                                    if (sbSHORHomeMonth.Length > 0)
                                        sbExtractList.Append(sbSHORHomeMonth.ToString());

                                    //4c:
                                    //Homeless_History_Yr | Only do if original submission or resubmission
                                    if (sbSHORHomeYear.Length > 0)
                                        sbExtractList.Append(sbSHORHomeYear.ToString());

                                    //4d:
                                    //Institutions | Only do if original submission or resubmission
                                    if (sbSHORFac.Length > 0)
                                        sbExtractList.Append(sbSHORFac.ToString());

                                    //4e:
                                    //Reasons | Only do if original submission or resubmission
                                    if (sbSHORReas.Length > 0)
                                        sbExtractList.Append(sbSHORReas.ToString());

                                    //4f new:
                                    //Disability | Only do if original submission or resubmission
                                    if (varMonthStart >= Convert.ToDateTime("01-Jul-2013"))
                                    {
                                        //Disability Count
                                        varDFcount = varDFcount + 3;

                                        sbExtractList.AppendLine("         <SP_Disability_Ind>");
                                        sbExtractList.AppendLine("            <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                        sbExtractList.AppendLine("            <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                        sbExtractList.AppendLine("            <Disability_life_area_activity>1</Disability_life_area_activity>");
                                        sbExtractList.AppendLine("            <Disability_need_for_assistance>" + dbDisSelf + "</Disability_need_for_assistance>");
                                        sbExtractList.AppendLine("         </SP_Disability_Ind>");

                                        sbExtractList.AppendLine("         <SP_Disability_Ind>");
                                        sbExtractList.AppendLine("            <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                        sbExtractList.AppendLine("            <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                        sbExtractList.AppendLine("            <Disability_life_area_activity>2</Disability_life_area_activity>");
                                        sbExtractList.AppendLine("            <Disability_need_for_assistance>" + dbDisMob + "</Disability_need_for_assistance>");
                                        sbExtractList.AppendLine("         </SP_Disability_Ind>");

                                        sbExtractList.AppendLine("         <SP_Disability_Ind>");
                                        sbExtractList.AppendLine("            <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                        sbExtractList.AppendLine("            <Support_Period_ID>" + varSuppId + "</Support_Period_ID>");
                                        sbExtractList.AppendLine("            <Disability_life_area_activity>3</Disability_life_area_activity>");
                                        sbExtractList.AppendLine("            <Disability_need_for_assistance>" + dbDisComm + "</Disability_need_for_assistance>");
                                        sbExtractList.AppendLine("         </SP_Disability_Ind>");
                                    }

                               // } //varSubmission != 2

                                //End 3:
                                sbExtractList.AppendLine("      </SP_Support_Period>");

                            } // Client Loop

                            //6:
                            //Turnaway
                            //Loop through turnaway records
                            foreach (var c in result10.Entities)
                            {
                                //varTest += c.Attributes["new_firstname"];

                                //foreach (KeyValuePair<String, Object> attribute in c.Attributes)
                                //{
                                //    varTest += (attribute.Key + ": " + attribute.Value + "\r\n");
                                //}

                                //foreach (KeyValuePair<String, String> value in c.FormattedValues)
                                //{
                                //    varTest += (value.Key + ": " + value.Value + "\r\n");
                                //}

                                // Reset stringbuilders
                                sbSHORTAR.Length = 0;
                                sbSHORTAQ.Length = 0;

                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (c.FormattedValues.Contains("new_agencyid"))
                                    dbStepAgencyId = c.FormattedValues["new_agencyid"];
                                else if (c.Attributes.Contains("new_agencyid"))
                                    dbStepAgencyId = c.Attributes["new_agencyid"].ToString();
                                else
                                    dbStepAgencyId = "";

                                if (c.FormattedValues.Contains("new_date"))
                                    dbStepDate = c.FormattedValues["new_date"];
                                else if (c.Attributes.Contains("new_date"))
                                    dbStepDate = c.Attributes["new_date"].ToString();
                                else
                                    dbStepDate = "";

                                // Convert date from American format to Australian format
                                dbStepDate = cleanDateAM(dbStepDate);

                                if (c.FormattedValues.Contains("new_dob"))
                                    dbStepDob = c.FormattedValues["new_dob"];
                                else if (c.Attributes.Contains("new_dob"))
                                    dbStepDob = c.Attributes["new_dob"].ToString();
                                else
                                    dbStepDob = "";

                                // Convert date from American format to Australian format
                                dbStepDob = cleanDateAM(dbStepDob);

                                if (c.FormattedValues.Contains("new_estimated"))
                                    dbStepEstimated = c.FormattedValues["new_estimated"];
                                else if (c.Attributes.Contains("new_estimated"))
                                    dbStepEstimated = c.Attributes["new_estimated"].ToString();
                                else
                                    dbStepEstimated = "";

                                if (c.FormattedValues.Contains("new_dobestd"))
                                    dbDobEstD = c.FormattedValues["new_dobestd"];
                                else if (c.Attributes.Contains("new_dobestd"))
                                    dbDobEstD = c.Attributes["new_dobestd"].ToString();
                                else
                                    dbDobEstD = "";

                                if (c.FormattedValues.Contains("new_dobestm"))
                                    dbDobEstM = c.FormattedValues["new_dobestm"];
                                else if (c.Attributes.Contains("new_dobestm"))
                                    dbDobEstM = c.Attributes["new_dobestm"].ToString();
                                else
                                    dbDobEstM = "";

                                if (c.FormattedValues.Contains("new_dobesty"))
                                    dbDobEstY = c.FormattedValues["new_dobesty"];
                                else if (c.Attributes.Contains("new_dobesty"))
                                    dbDobEstY = c.Attributes["new_dobesty"].ToString();
                                else
                                    dbDobEstY = "";

                                // Get dobflag in new format (day, month, year seperate) or as single checkbox
                                varDobFlag = "AAA";
                                if (string.IsNullOrEmpty(dbDobEstD) == false || string.IsNullOrEmpty(dbDobEstM) == false || string.IsNullOrEmpty(dbDobEstY) == false)
                                {
                                    if (string.IsNullOrEmpty(dbDobEstD) == false)
                                        varDobFlag = dbDobEstD.Substring(0, 1) + varDobFlag.Substring(1, 2);
                                    if (string.IsNullOrEmpty(dbDobEstM) == false)
                                        varDobFlag = varDobFlag.Substring(0, 1) + dbDobEstM.Substring(0, 1) + varDobFlag.Substring(2, 1);
                                    if (string.IsNullOrEmpty(dbDobEstY) == false)
                                        varDobFlag = varDobFlag.Substring(0, 2) + dbDobEstY.Substring(0, 1);
                                }
                                else if (dbStepEstimated == "Yes")
                                {
                                    varDobFlag = "EEE";
                                }

                                dbStepEstimated = varDobFlag;

                                if (c.FormattedValues.Contains("new_firstname"))
                                    dbStepFirstName = c.FormattedValues["new_firstname"];
                                else if (c.Attributes.Contains("new_firstname"))
                                    dbStepFirstName = c.Attributes["new_firstname"].ToString();
                                else
                                    dbStepFirstName = "";

                                if (c.FormattedValues.Contains("new_firstserv"))
                                    dbStepFirstServ = c.FormattedValues["new_firstserv"];
                                else if (c.Attributes.Contains("new_firstserv"))
                                    dbStepFirstServ = c.Attributes["new_firstserv"].ToString();
                                else
                                    dbStepFirstServ = "";

                                if (c.FormattedValues.Contains("new_sex"))
                                    dbStepGender = c.FormattedValues["new_sex"];
                                else if (c.Attributes.Contains("new_sex"))
                                    dbStepGender = c.Attributes["new_sex"].ToString();
                                else
                                    dbStepGender = "";

                                // Get entity id for client
                                if (c.Attributes.Contains("new_ispalm"))
                                {
                                    getEntity = (EntityReference)c.Attributes["new_ispalm"];
                                    dbStepIsPalm = getEntity.Id.ToString();
                                }
                                else if (c.FormattedValues.Contains("new_ispalm"))
                                    dbStepIsPalm = c.FormattedValues["new_ispalm"];
                                else
                                    dbStepIsPalm = "";

                                // Get entity id for turnaway
                                if (c.Attributes.Contains("new_isstep"))
                                {
                                    getEntity = (EntityReference)c.Attributes["new_isstep"];
                                    dbStepIsStep = getEntity.Id.ToString();
                                }
                                else if (c.FormattedValues.Contains("new_isstep"))
                                    dbStepIsStep = c.FormattedValues["new_isstep"];
                                else
                                    dbStepIsStep = "";

                                if (c.FormattedValues.Contains("new_name"))
                                    dbStepName = c.FormattedValues["new_name"];
                                else if (c.Attributes.Contains("new_name"))
                                    dbStepName = c.Attributes["new_name"].ToString();
                                else
                                    dbStepName = "";

                                if (c.FormattedValues.Contains("new_relationship"))
                                    dbStepRelationship = c.FormattedValues["new_relationship"];
                                else if (c.Attributes.Contains("new_relationship"))
                                    dbStepRelationship = c.Attributes["new_relationship"].ToString();
                                else
                                    dbStepRelationship = "";

                                if (c.FormattedValues.Contains("new_relother"))
                                    dbStepRelOther = c.FormattedValues["new_relother"];
                                else if (c.Attributes.Contains("new_relother"))
                                    dbStepRelOther = c.Attributes["new_relother"].ToString();
                                else
                                    dbStepRelOther = "";

                                if (c.FormattedValues.Contains("new_surname"))
                                    dbStepSurname = c.FormattedValues["new_surname"];
                                else if (c.Attributes.Contains("new_surname"))
                                    dbStepSurname = c.Attributes["new_surname"].ToString();
                                else
                                    dbStepSurname = "";

                                if (c.FormattedValues.Contains("new_turnassist"))
                                    dbStepTurnAssist = c.FormattedValues["new_turnassist"];
                                else if (c.Attributes.Contains("new_turnassist"))
                                    dbStepTurnAssist = c.Attributes["new_turnassist"].ToString();
                                else
                                    dbStepTurnAssist = "";

                                if (c.FormattedValues.Contains("new_turnreason"))
                                    dbStepTurnReason = c.FormattedValues["new_turnreason"];
                                else if (c.Attributes.Contains("new_turnreason"))
                                    dbStepTurnReason = c.Attributes["new_turnreason"].ToString();
                                else
                                    dbStepTurnReason = "";

                                // Wrap values in asterisks for easier comparisons
                                dbStepTurnReason = getMult(dbStepTurnReason);

                                if (c.FormattedValues.Contains("new_turnreasoth"))
                                    dbStepTurnReasOth = c.FormattedValues["new_turnreasoth"];
                                else if (c.Attributes.Contains("new_turnreasoth"))
                                    dbStepTurnReasOth = c.Attributes["new_turnreasoth"].ToString();
                                else
                                    dbStepTurnReasOth = "";

                                if (c.FormattedValues.Contains("new_turnrequest"))
                                    dbStepTurnRequest = c.FormattedValues["new_turnrequest"];
                                else if (c.Attributes.Contains("new_turnrequest"))
                                    dbStepTurnRequest = c.Attributes["new_turnrequest"].ToString();
                                else
                                    dbStepTurnRequest = "";

                                // Wrap values in asterisks for easier comparisons
                                dbStepTurnRequest = getMult(dbStepTurnRequest);

                                if (c.FormattedValues.Contains("new_turnurg"))
                                    dbStepTurnUrg = c.FormattedValues["new_turnurg"];
                                else if (c.Attributes.Contains("new_turnurg"))
                                    dbStepTurnUrg = c.Attributes["new_turnurg"].ToString();
                                else
                                    dbStepTurnUrg = "";

                                if (c.FormattedValues.Contains("new_palmstepid"))
                                    dbPalmStepId = c.FormattedValues["new_palmstepid"];
                                else if (c.Attributes.Contains("new_palmstepid"))
                                    dbPalmStepId = c.Attributes["new_palmstepid"].ToString();
                                else
                                    dbPalmStepId = "";

                                // Set to self if the puh is the same as the turnaway id
                                if (dbStepRelationship.ToLower() == "self")
                                    dbStepIsStep = dbPalmStepId;

                                // Remove dashes from id for extract
                                varStepId = "";
                                if (String.IsNullOrEmpty(dbPalmStepId) == false)
                                    varStepId = dbPalmStepId.Replace("-", "");
                                varPuhId = "";
                                if (String.IsNullOrEmpty(dbStepIsStep) == false)
                                    varPuhId = dbStepIsStep.Replace("-", "");

                                // Cannot use value 5 before 1 July 2019
                                if (Convert.ToDateTime(dbAssessed) < Convert.ToDateTime("1-Jul-2019") && String.IsNullOrEmpty(dbStepTurnRequest) == false)
                                    dbStepTurnRequest = dbStepTurnRequest.Replace("Assistance for family and domestic violence", "3");

                                //[SHOR](103) | Services requested
                                if (String.IsNullOrEmpty(dbStepTurnRequest) == true)
                                    dbStepTurnRequest = "*3*";

                                //[SHOR](106) | Reason for turnaway
                                if (String.IsNullOrEmpty(dbStepTurnReason) == true)
                                    dbStepTurnReason = "*11*";

                                //Get the values from the drop down tables
                                foreach (var d in result2.Entities)
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

                                    if (d.FormattedValues.Contains("new_shor"))
                                        varSHOR = d.FormattedValues["new_shor"];
                                    else if (d.Attributes.Contains("new_shor"))
                                        varSHOR = d.Attributes["new_shor"].ToString();
                                    else
                                        varSHOR = "";

                                    varSHOR = cleanString(varSHOR, "number");
                                    if (String.IsNullOrEmpty(varSHOR) == true)
                                        varSHOR = "0";

                                    // Set the turnaway urgency to a numeric value
                                    if (varType == "turnawayurg" && varDesc == dbStepTurnUrg)
                                        dbStepTurnUrg = varSHOR + "";

                                    // Set the turnaway relationship to a numeric value
                                    if (varType == "puhrship" && varDesc == dbStepRelationship)
                                        dbStepRelationship = varSHOR + "";

                                    //Create the turnaway reason part of the extract
                                    if (varType == "turnawayreas")
                                    {
                                        // If the reason is not null and contains the number / description, then add to the extract
                                        if (String.IsNullOrEmpty(dbStepTurnReason) == false)
                                        {
                                            if (dbStepTurnReason.IndexOf("*" + varSHOR + "*") > -1 || dbStepTurnReason.IndexOf("*" + varDesc + "*") > -1)
                                            {
                                                //Turnaway Reason Count
                                                varTRcount++;

                                                sbSHORTAR.AppendLine("         <Turnaway_Reason>");
                                                sbSHORTAR.AppendLine("            <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                sbSHORTAR.AppendLine("            <Turnaway_ID>" + varStepId + "</Turnaway_ID>");
                                                sbSHORTAR.AppendLine("            <Reason_Service_Not_Provided>" + varSHOR + "</Reason_Service_Not_Provided>");
                                                sbSHORTAR.AppendLine("         </Turnaway_Reason>");

                                            }
                                        }

                                    }

                                    //Create the turnaway request part of the extract
                                    if (varType == "turnawayreq")
                                    {
                                        // If the requested service is not null and contains the number / description, then add to the extract
                                        if (String.IsNullOrEmpty(dbStepTurnRequest) == false)
                                        {
                                            if (dbStepTurnRequest.IndexOf("*" + varSHOR + "*") > -1 || dbStepTurnRequest.IndexOf("*" + varDesc + "*") > -1)
                                            {
                                                //FIX
                                                if (varSHOR != "5" || (varSHOR == "5" && Convert.ToDateTime(dbAssessed) >= Convert.ToDateTime("1-Jul-2019")))
                                                {
                                                    //Turnaway Service Count
                                                    varTScount++;

                                                    sbSHORTAQ.AppendLine("         <Turnaway_Service_Request>");
                                                    sbSHORTAQ.AppendLine("            <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                                    sbSHORTAQ.AppendLine("            <Turnaway_ID>" + varStepId + "</Turnaway_ID>");
                                                    sbSHORTAQ.AppendLine("            <Service_Requested>" + varSHOR + "</Service_Requested>");
                                                    sbSHORTAQ.AppendLine("         </Turnaway_Service_Request>");
                                                }

                                            }
                                        }

                                    }

                                } // Drop List Loop

                                //[SHOR](88) | Turnaway ID
                                if (String.IsNullOrEmpty(dbPalmStepId) == true)
                                {
                                    dbPalmStepId = "0";
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " " + dbStartDate + " - Must have Turn away ID");
                                }
                                else if (dbPalmStepId.Length > 50)
                                {
                                    dbPalmStepId = "0";
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " " + dbStartDate + " - Turn away ID greater than 50 characters");
                                }


                                //[SHOR](89) | Letters of Family Name
                                varSurname = dbStepSurname.ToUpper() + "22222";
                                varSurname = cleanString(varSurname, "slk");
                                varSurname = varSurname.Substring(1, 1) + varSurname.Substring(2, 1) + varSurname.Substring(4, 1);

                                if (varSurname == "222")
                                    varSurname = "999";

                                //[SHOR](90) | Letters of Given Name
                                varFirstName = dbStepFirstName.ToUpper() + "22222";
                                varFirstName = cleanString(varFirstName, "slk");

                                varFirstName = varFirstName.Substring(1, 1) + varFirstName.Substring(2, 1);

                                if (varFirstName == "22")
                                    varFirstName = "99";

                                //[SHOR](91) | Sex
                                if (dbStepGender == "Male")
                                    dbStepGender = "1";
                                else if (dbStepGender == "Other")
                                {
                                    if (Convert.ToDateTime(dbStepDate) < Convert.ToDateTime("1-Jul-2019") || varMonthStart < Convert.ToDateTime("1-Jul-2019"))
                                        dbStepGender = "2";
                                    else
                                        dbStepGender = "3";
                                }
                                else
                                    dbStepGender = "2";

                                //[SHOR](93) | Date of Birth Estimate
                                if (dbStepEstimated == "estimated")
                                    dbStepEstimated = "EEE";
                                else if (dbStepEstimated == "not estimated")
                                    dbStepEstimated = "AAA";

                                dbStepEstimated = dbStepEstimated + "AAA";
                                dbStepEstimated = dbStepEstimated.Substring(0, 3);

                                if (dbStepEstimated.Substring(0, 1) != "A" && dbStepEstimated.Substring(0, 1) != "E" && dbStepEstimated.Substring(0, 1) != "U")
                                    dbStepEstimated = "AAA";
                                if (dbStepEstimated.Substring(1, 1) != "A" && dbStepEstimated.Substring(1, 1) != "E" && dbStepEstimated.Substring(1, 1) != "U")
                                    dbStepEstimated = "AAA";
                                if (dbStepEstimated.Substring(2, 1) != "A" && dbStepEstimated.Substring(2, 1) != "E" && dbStepEstimated.Substring(2, 1) != "U")
                                    dbStepEstimated = "AAA";

                                //[SHOR](92) | Date of Birth
                                if (String.IsNullOrEmpty(dbStepDob) == false)
                                {
                                    if (Convert.ToDateTime(dbStepDob).Year >= 1880)
                                    {
                                        if (Convert.ToDateTime(dbStepDob) < Convert.ToDateTime(DateTime.Now.AddYears(-116)))
                                        {
                                            dbStepDob = "1-Jan-1880";
                                            dbStepEstimated = "UUU";
                                        }
                                    }
                                    else
                                    {
                                        dbStepDob = "1-Jan-1880";
                                        dbStepEstimated = "UUU";
                                        sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " " + dbStartDate + " - DOB");
                                    }
                                }
                                else
                                {
                                    dbStepDob = "1-Jan-1880";
                                    dbStepEstimated = "UUU";
                                }

                                if (dbStepEstimated == "UUU")
                                    dbStepDob = "1-Jan-1880";

                                // Get age of turnaway
                                varAge = varMonthEnd.Year - Convert.ToDateTime(dbStepDob).Year;
                                if (varMonthEnd.Month < Convert.ToDateTime(dbStepDob).Month || (varMonthEnd.Month == Convert.ToDateTime(dbStepDob).Month && varMonthEnd.Day < Convert.ToDateTime(dbStepDob).Day))
                                    varAge = varAge - 1;

                                //[SHOR](94) | Assistance Date
                                if (String.IsNullOrEmpty(dbStepDate) == false)
                                {

                                    if (Convert.ToDateTime(dbStepDate).Year < 1880)
                                    {
                                        dbStepDate = "";
                                        sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " " + dbStartDate + " - Assessed Date");
                                    }

                                }
                                else
                                {
                                    dbStepDate = "";
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " " + dbStartDate + " - Assessed Date");
                                }

                                //[SHOR](95) | Presenting Unit Head ID
                                if (String.IsNullOrEmpty(dbStepIsStep) == true)
                                {
                                    dbStepIsStep = "0";
                                    dbStepRelOther = "";
                                    dbStepRelationship = "1";
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " " + dbStartDate + " - No turnaway presenting unit head id");
                                }
                                else if (dbStepIsStep.Length > 50)
                                {
                                    dbStepIsStep = "0";
                                    dbStepRelOther = "";
                                    dbStepRelationship = "1";
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " " + dbStartDate + " - Turnaway presenting unit head ID greater than 50 characters");
                                }

                                //[SHOR](96) | Relationship to Presenting Unit Head
                                if (!Int32.TryParse(dbStepRelationship, out varCheckInt))
                                    dbStepRelationship = "99";

                                //[SHOR](97) | Relationship to Presenting Unit Head Other
                                if (dbStepRelationship == "12" || dbStepRelationship == "15")
                                {
                                    if (String.IsNullOrEmpty(dbStepRelOther) == true)
                                        dbStepRelOther = "No Value";
                                    else if (dbStepRelOther.Length > 100)
                                        dbStepRelOther = "No Value";

                                }
                                else
                                    dbStepRelOther = "";

                                //[SHOR](98) | Count in presenting unit
                                varPUHcount = 0;

                                // Get count in presenting unit
                                foreach (var c2 in result10.Entities)
                                {
                                    if (c2.Attributes.Contains("new_isstep"))
                                    {
                                        getEntity = (EntityReference)c2.Attributes["new_isstep"];
                                        dbStepIsStep2 = getEntity.Id.ToString();
                                    }
                                    else if (c2.FormattedValues.Contains("new_isstep"))
                                        dbStepIsStep2 = c2.FormattedValues["new_isstep"];
                                    else
                                        dbStepIsStep2 = "";

                                    if (c2.FormattedValues.Contains("new_palmstepid"))
                                        dbPalmStepId2 = c2.FormattedValues["new_palmstepid"];
                                    else if (c2.Attributes.Contains("new_palmstepid"))
                                        dbPalmStepId2 = c2.Attributes["new_palmstepid"].ToString();
                                    else
                                        dbPalmStepId2 = "";

                                    if (c2.FormattedValues.Contains("new_relationship"))
                                        dbStepRelationship2 = c2.FormattedValues["new_relationship"];
                                    else if (c2.Attributes.Contains("new_relationship"))
                                        dbStepRelationship2 = c2.Attributes["new_relationship"].ToString();
                                    else
                                        dbStepRelationship2 = "";

                                    if (dbStepRelationship2.ToLower() == "self")
                                        dbStepIsStep2 = dbPalmStepId2;

                                    //Get the PUH count
                                    if (dbStepIsStep == dbStepIsStep2)
                                        varPUHcount++;
                                } //puh Loop

                                //[SHOR](99) | Urgency of turnaway
                                if (!Int32.TryParse(dbStepTurnUrg, out varCheckInt))
                                    dbStepTurnUrg = "99";

                                //[SHOR](100) | Was it first service for day?
                                if (dbStepFirstServ == "Yes")
                                    dbStepFirstServ = "1";
                                else if (dbStepFirstServ == "No")
                                    dbStepFirstServ = "2";
                                else
                                    dbStepFirstServ = "99";


                                //Validation rules

                                //C04.006.02 | Assessed date cant be greater than DOB
                                if (String.IsNullOrEmpty(dbStepDob) == false && String.IsNullOrEmpty(dbStepDate) == false)
                                {
                                    if (Convert.ToDateTime(dbStepDob) > Convert.ToDateTime(dbStepDate))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " - date of birth is greater than than assessed date");
                                }

                                //C04.009.02 | If relationship to PUH is self and unit count = 1 then IDs must be the same
                                if (dbStepIsStep == "1" && varPUHcount == 1 && dbStepIsStep != dbPalmStepId)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " - Turnaway Presenting unit id is different to own id but relationship is self");

                                //C04.009.03 | If relationship to PUH not self and unit count > 1 then IDs cant be the same
                                if (dbStepIsStep != "1" && varPUHcount > 1 && dbStepIsStep == dbPalmStepId)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " - Turnaway Presenting unit is not self but has same id as self");

                                //C04.010.02 | If IDs are the same and unit count = 1 then relationship needs to be self
                                if (dbStepIsStep == dbPalmStepId && varPUHcount == 1)
                                {
                                    dbStepRelOther = "";
                                    dbStepRelationship = "1";
                                }

                                //C04.010.03 | If IDs are not the same and unit count > 1 then relationship cant be self
                                if (dbStepIsStep != dbPalmStepId && varPUHcount > 1 && dbStepRelationship == "1")
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " - Turnaway Presenting unit is not self and unit count greater than 1 but relationship is self");

                                //C04.010.04 | Can't be grandparent if less than 18 (can be grand parent if 19 though)
                                if (dbStepRelationship == "10" && varAge < 18)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " - Turnaway Relationship is grandparents but person us under 18");

                                //C04.012.02 / W07.019.01 | Unit count cant be 0 or greater than 15
                                if (varPUHcount == 0)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " - Turnaway Presenting unit head count is 0");
                                else if (varPUHcount > 15)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " - Turnaway Presenting unit head count is greater than 15");

                                //C04.012.03 | Cant have unit count > 1 if IDs are the same
                                if (varPUHcount > 1 && dbStepIsStep == dbPalmStepId)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " - Turnaway Presenting unit head count is greater than 1 but id is same");

                                //C04.012.04 | If relationship not self, can't have unit count of 1
                                if (dbStepRelationship != "1" && varPUHcount == 1)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbStepName + ") " + dbStepFirstName + " " + dbStepSurname + " - Turnaway Presenting unit head count is 1 but relationship is not self");

                                //Turnaway Count
                                varTAcount++;

                                // Append turnaway to extract
                                sbExtractList.AppendLine("      <Turnaway>");
                                sbExtractList.AppendLine("         <Organisation_ID>" + varAgencyId + "</Organisation_ID>");
                                sbExtractList.AppendLine("         <Turnaway_ID>" + varStepId + "</Turnaway_ID>"); //How to handle this number?
                                sbExtractList.AppendLine("         <Letters_Of_Family_Name>" + varSurname + "</Letters_Of_Family_Name>");
                                sbExtractList.AppendLine("         <Letters_Of_Given_Name>" + varFirstName + "</Letters_Of_Given_Name>");
                                sbExtractList.AppendLine("         <Sex>" + dbStepGender + "</Sex>");
                                sbExtractList.AppendLine("         <Date_Of_Birth>" + cleanDateS(Convert.ToDateTime(dbStepDob)) + "</Date_Of_Birth>");
                                sbExtractList.AppendLine("         <Date_Of_Birth_Accuracy_Ind>" + dbStepEstimated + "</Date_Of_Birth_Accuracy_Ind>");
                                sbExtractList.AppendLine("         <Assistance_Request_Date>" + cleanDateS(Convert.ToDateTime(dbStepDate)) + "</Assistance_Request_Date>");
                                sbExtractList.AppendLine("         <PULK_Turnaway_ID>" + varPuhId + "</PULK_Turnaway_ID>"); //Unique to turnaway head???
                                sbExtractList.AppendLine("         <Relationship_To_PUH_TA>" + dbStepRelationship + "</Relationship_To_PUH_TA>");

                                if (String.IsNullOrEmpty(dbStepRelOther) == false)
                                    sbExtractList.AppendLine("         <Relationship_To_PUH_TA_Other>" + dbStepRelOther + "</Relationship_To_PUH_TA_Other>");
                                else
                                    sbExtractList.AppendLine("         <Relationship_To_PUH_TA_Other xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />");

                                sbExtractList.AppendLine("         <Count_In_Presenting_Unit>" + varPUHcount + "</Count_In_Presenting_Unit>");
                                sbExtractList.AppendLine("         <Urgency_Of_Request>" + dbStepTurnUrg + "</Urgency_Of_Request>");
                                sbExtractList.AppendLine("         <First_Service_Request_Ind>" + dbStepFirstServ + "</First_Service_Request_Ind>");

                                //6a
                                //Reason turned away

                                if (sbSHORTAR.Length > 0)
                                    sbExtractList.Append(sbSHORTAR.ToString());

                                //6b
                                //Service denied

                                if (sbSHORTAQ.Length > 0)
                                    sbExtractList.Append(sbSHORTAQ.ToString());


                                //End 6
                                sbExtractList.AppendLine("      </Turnaway>");

                                sbClientList.AppendLine("UNASSIST," + varStepId + "," + dbStepName + ", " + dbStepFirstName + "," + dbStepSurname + "," + dbStepDate + ",");

                            } // Turnaway Loop


                            // Create SHS extract
                            sbHeaderList.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");
                            //1:
                            //Basic extract information
                            sbHeaderList.AppendLine("<Extract>");
                            sbHeaderList.AppendLine("   <Collection_Period>" + varPrintMonth + varMonthStart.Year + "</Collection_Period>"); //[SHOR](1)
                            sbHeaderList.AppendLine("   <Software_Product>HHS Connect</Software_Product>"); //[SHOR](2) 'Must be certified
                            sbHeaderList.AppendLine("   <Software_Version>1.1</Software_Version>"); //[SHOR](3) 'Must be certified
                            sbHeaderList.AppendLine("   <Extract_Agency_Cnt>1</Extract_Agency_Cnt>"); //Not auto | Count <Extract_Agency> [SHOR](4)

                            //2:
                            //Summary and statistics
                            sbHeaderList.AppendLine("   <Extract_Agency>");
                            sbHeaderList.AppendLine("      <Organisation_ID>" + varAgencyId + "</Organisation_ID>"); //[SHOR](5)
                            sbHeaderList.AppendLine("      <Organisation_Name>" + varAgencyName + "</Organisation_Name>"); //[SHOR](6)
                            sbHeaderList.AppendLine("      <Turnaway_Cnt>" + varTAcount + "</Turnaway_Cnt>"); //[SHOR](7)
                            sbHeaderList.AppendLine("      <Turnaway_Service_Request_Cnt>" + varTScount + "</Turnaway_Service_Request_Cnt>"); //[SHOR](8)
                            sbHeaderList.AppendLine("      <Turnaway_Reason_Cnt>" + varTRcount + "</Turnaway_Reason_Cnt>"); //[SHOR](9)
                            sbHeaderList.AppendLine("      <SP_Support_Period_Cnt>" + varSPcount + "</SP_Support_Period_Cnt>"); //[SHOR](10)
                            sbHeaderList.AppendLine("      <SP_Institutions_Cnt>" + varINcount + "</SP_Institutions_Cnt>"); //[SHOR](11)
                            sbHeaderList.AppendLine("      <SP_Reasons_Cnt>" + varREcount + "</SP_Reasons_Cnt>"); //[SHOR](12)
                            sbHeaderList.AppendLine("      <SP_Homeless_History_Mth_Cnt>" + varHMcount + "</SP_Homeless_History_Mth_Cnt>"); //[SHOR](13)
                            sbHeaderList.AppendLine("      <SP_Homeless_History_Yr_Cnt>" + varHYcount + "</SP_Homeless_History_Yr_Cnt>"); //[SHOR](14)

                            if (Convert.ToDateTime(varMonthStart) >= Convert.ToDateTime("01-Jul-2013"))
                                sbHeaderList.AppendLine("      <SP_Disability_Ind_Cnt>" + varDFcount + "</SP_Disability_Ind_Cnt>"); //NEW Disability flag

                            sbHeaderList.AppendLine("      <SP_CP_Collection_Period_Cnt>" + varCPcount + "</SP_CP_Collection_Period_Cnt>"); //[SHOR](15)
                            sbHeaderList.AppendLine("      <SP_CP_Support_Services_Cnt>" + varSScount + "</SP_CP_Support_Services_Cnt>"); //[SHOR](16)
                            sbHeaderList.AppendLine("      <SP_CP_Accomm_Periods_Cnt>" + varAPcount + "</SP_CP_Accomm_Periods_Cnt>"); //[SHOR](17)
                            sbHeaderList.AppendLine("      <SP_CP_Financial_Support_Cnt>" + varFScount + "</SP_CP_Financial_Support_Cnt>"); //[SHOR](18)
                            sbHeaderList.AppendLine("      <SP_CP_Homeless_History_Cnt>" + varHHcount + "</SP_CP_Homeless_History_Cnt>"); //[SHOR](19)

                            if (sbExtractList.Length > 0)
                                sbHeaderList.Append(sbExtractList.ToString());

                            sbHeaderList.AppendLine("   </Extract_Agency>");
                            sbHeaderList.AppendLine("</Extract>");


                            // Create note against current Palm Go SHOR record and add attachment
                            string strMessage = sbHeaderList.ToString();
                            byte[] filename = Encoding.ASCII.GetBytes(strMessage);
                            string encodedData = System.Convert.ToBase64String(filename);
                            Entity Annotation = new Entity("annotation");
                            Annotation.Attributes["objectid"] = new EntityReference("new_palmgoshor", varShorID);
                            Annotation.Attributes["objecttypecode"] = "new_palmgoshor";
                            Annotation.Attributes["subject"] = "SHS Extract";
                            Annotation.Attributes["documentbody"] = encodedData;
                            Annotation.Attributes["mimetype"] = @"text / plain";
                            Annotation.Attributes["notetext"] = "SHS Extract for " + varMonthStart.Year + varPrintMonth;
                            Annotation.Attributes["filename"] = "shor_extract_" + varAgencyId + "_" + varMonthStart.Year + varPrintMonth + ".xml";
                            _service.Create(Annotation);

                            // Add the client list
                            byte[] filename3 = Encoding.ASCII.GetBytes(sbClientList.ToString());
                            string encodedData3 = System.Convert.ToBase64String(filename3);
                            Entity Annotation3 = new Entity("annotation");
                            Annotation.Attributes["objectid"] = new EntityReference("new_palmgoshor", varShorID);
                            Annotation.Attributes["objecttypecode"] = "new_palmgoshor";
                            Annotation.Attributes["subject"] = "SHS Extract";
                            Annotation.Attributes["documentbody"] = encodedData3;
                            Annotation.Attributes["mimetype"] = @"text/csv";
                            Annotation.Attributes["notetext"] = "SHS Client List for " + varMonthStart.Year + varPrintMonth;
                            Annotation.Attributes["filename"] = "shor_clients_" + varAgencyId + "_" + varMonthStart.Year + varPrintMonth + ".csv";
                            _service.Create(Annotation);


                            if (sbErrorList.Length > 0)
                            {
                                // If there is an error, create note against current Palm Go SHOR record and add attachment
                                byte[] filename2 = Encoding.ASCII.GetBytes(sbErrorList.ToString());
                                string encodedData2 = System.Convert.ToBase64String(filename2);
                                Entity Annotation2 = new Entity("annotation");
                                Annotation2.Attributes["objectid"] = new EntityReference("new_palmgoshor", varShorID);
                                Annotation2.Attributes["objecttypecode"] = "new_palmgoshor";
                                Annotation2.Attributes["subject"] = "SHS Extract";
                                Annotation2.Attributes["documentbody"] = encodedData2;
                                Annotation2.Attributes["mimetype"] = @"text / plain";
                                Annotation2.Attributes["notetext"] = "SHS errors and warnings for " + varAgencyId + "_" + varMonthStart.Year + varPrintMonth;
                                Annotation2.Attributes["filename"] = "Errors.txt";
                                _service.Create(Annotation2);
                            }

                            //throw new InvalidPluginExecutionException("The plugin is working:\r\n" + varTest);
                        }
                        catch (Exception e)
                        {
                            throw new InvalidPluginExecutionException("An error occured:\r\n" + e);
                        }
                    }

                    
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in the FollowupPlugin plug-in.", ex);
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

        //Remove characters from a string
        public string removeString(string clean, string thetype)
        {
            string varCharNotAllowed = ""; //Characters allower
            string temp = ""; //Temporary string for removing illegal characters

            if (String.IsNullOrEmpty(clean) == false)
            {
                clean = clean.Trim();
                temp = clean;
            }

            //Replace <> characters with html code
            if (String.IsNullOrEmpty(clean) == false)
            {
                clean = clean.Replace("<", "&lt;");
                clean = clean.Replace(">", "&gt;");
            }

            if (thetype == "html")
                varCharNotAllowed = "<>"; //Characters not allowed
            else if (thetype == "drop")
                varCharNotAllowed = "<\">"; //Characters not allowed
            else if (thetype == "file")
                varCharNotAllowed = "<>*?\"|"; //Characters not allowed
            else if (thetype == "link")
                varCharNotAllowed = "<>?\"|&="; //Characters not allowed
            else if (thetype == "image")
                varCharNotAllowed = "\":<>"; //Characters not allowed
            else if (thetype == "coconut")
            {

                if (String.IsNullOrEmpty(clean) == false)
                {
                    clean = clean.Replace("*,", "*----");
                    clean = clean.Replace(",", "**");
                    clean = clean.Replace("\r", "~~~~");
                    clean = clean.Replace("\n", "++++");
                    clean = clean.Replace("\"", "'");
                }

            }
            else if (thetype == "decoconut")
            {

                if (String.IsNullOrEmpty(clean) == false)
                {
                    clean = clean.Replace("*----", "*,");
                    clean = clean.Replace("**", ",");
                    clean = clean.Replace("----", ",");
                    clean = clean.Replace("~~~~", "\r");
                    clean = clean.Replace("++++", "\n");
                }

            }
            else
                varCharNotAllowed = "<>\"'"; //Characters not allowed

            //Set a temporary string to the value of the string passed
            temp = clean;

            if (String.IsNullOrEmpty(clean) == false)
            {

                //Loop through each character in the forbidden character string
                for (int k = 0; k < varCharNotAllowed.Length; k++)
                {
                    //If the next character is in the set of characters, replace it with ~
                    if (clean.IndexOf(varCharNotAllowed[k]) > -1 && varCharNotAllowed[k].ToString() != "~")
                        temp = temp.Replace(varCharNotAllowed[k].ToString(), "~");
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

            //if (!DateTime.TryParse(clean, out varCheckDate))
            //clean = DateTime.Now.Day + "-" + DateTime.Now.ToString("MMM") + "-" + DateTime.Now.Year;

            if (!DateTime.TryParse(clean, out varCheckDate))
                clean = "";

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

        //Date format for SHOR: 01062013 - BIRTHDATE
        public string cleanDateSBirthDate(DateTime getdate, String acc)
        {
            string varDay = "";
            string varMonth = "";
            string varYear = "";

            if (getdate.Day < 10)
                varDay = "0" + getdate.Day;
            else
                varDay = getdate.Day + "";

            if (getdate.Month < 10)
                varMonth = "0" + getdate.Month;
            else
                varMonth = getdate.Month + "";

            varYear = "" + getdate.Year;

            acc = acc.ToLower();

            if(acc != "aaa")
            {
                char[] charArr = new char[] { };

                charArr = acc.ToCharArray();
                if (charArr[0] != 'a' && charArr[0] != 'e')
                    varDay = "01";
                if (charArr[1] != 'a' && charArr[1] != 'e')
                    varMonth = "01";
                if (charArr[2] != 'a' && charArr[2] != 'e')
                    varYear = "1880";
            }

            string clean = varDay + varMonth + varYear;
            return clean;
        }

        //Date format for SHOR: 01062013
        public string cleanDateSLK(DateTime getdate)
        {
            string varDay = "";
            string varMonth = "";
            string clean = "";

            if (getdate.Day < 10)
                varDay = "0" + getdate.Day;
            else
                varDay = getdate.Day + "";

            if (getdate.Month < 10)
                varMonth = "0" + getdate.Month;
            else
                varMonth = getdate.Month + "";

            if (getdate.Hour >= 12)
            {

                if (getdate.Hour == 12 && getdate.Minute < 10)
                    clean = "120" + getdate.Minute + "pm";
                else if (getdate.Hour == 12)
                    clean = "12" + getdate.Minute + "pm";
                else if (getdate.Hour == 24 && getdate.Minute < 10)
                    clean = "120" + getdate.Minute + "am";
                else if (getdate.Hour == 24)
                    clean = "12" + getdate.Minute + "am";
                else if (getdate.Minute < 10)
                    clean = (getdate.Hour - 12) + "0" + getdate.Minute + "pm";
                else
                    clean = (getdate.Hour - 12) + "" + getdate.Minute + "pm";
            }
            else
            {
                if (getdate.Minute < 10)
                    clean = getdate.Hour + "0" + getdate.Minute + "am";
                else
                    clean = getdate.Hour + "" + getdate.Minute + "am";
            }

            clean = varDay + varMonth + getdate.Year + "_" + clean;
            return clean;
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


