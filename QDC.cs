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
	public class goQDC: IPlugin
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
                // if (entity.LogicalName != "new_palmgoqdc")
                //    return;

                try
                {
                    // Fixes:
                    // Remove ampersand from comments
                    // Compare values to Lower

                    // Global variables
                    string varDescription = ""; // Description field on form
                    bool varCreateExtract = false; // Create Extract field on form
                    int varYear = 0; // Year from form
                    int varQuarter = 0; // Quarter from form
                    string varQuarterText = ""; // Text for quarter
                    int varQuarterFull = 0; // Quarter expressed as optionset id
                    Guid varQdcID = new Guid(); // GUID for palm go qdc record
                    StringBuilder sbHeaderList = new StringBuilder(); // String builder for extract header
                    StringBuilder sbClientList = new StringBuilder(); // String builder for extract client data
                    StringBuilder sbErrorList = new StringBuilder(); // String builder for error list
                    string varFileName = ""; // Extract file name
                    string varFileName2 = ""; // Error log file name
                    DateTime varStartDate = new DateTime(); // Start date of extract
                    DateTime varEndDate = new DateTime(); // End date of extract
                    DateTime varCurrentDate = DateTime.Now; // Current date for comparisons
                    int varCheckInt = 0; // Used to see if data is valid integer
                    double varCheckDouble = 0;  // Used to see if data is valid double
                    DateTime varCheckDate = new DateTime();
                    int varQDCID = 0; // Used to see if data is valid date
                    int varAgencyValue = 0; // Value of agency type
                    EntityReference getEntity; // Object to get entity details

                    string varTest = ""; // Used for debug

                    // Only do this if the entity is the Palm Go QDC entity
                    if (entity.LogicalName == "new_palmgoqdc")
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

                        // Get info for current QDC record
                        varDescription = entity.GetAttributeValue<string>("new_description");
                        varCreateExtract = entity.GetAttributeValue<bool>("new_createextract");
                        varYear = entity.GetAttributeValue<int>("new_year");
                        varQuarter = entity.GetAttributeValue<OptionSetValue>("new_quarter").Value;
                        varQdcID = entity.Id;
                        varQuarterText = entity.FormattedValues["new_quarter"];

                        // Assign Quarter Full to the Option Set value, then subtract 100000000 to get the normal quarter value
                        varQuarterFull = varQuarter;
                        varQuarter = varQuarter - 100000000;

                        //Get the start and end date of period
                        if (varQuarter == 3)
                        {
                            varStartDate = Convert.ToDateTime("1-Jan-" + varYear);
                            varEndDate = Convert.ToDateTime("31-Mar-" + varYear);
                        }
                        else if (varQuarter == 4)
                        {
                            varStartDate = Convert.ToDateTime("1-Apr-" + varYear);
                            varEndDate = Convert.ToDateTime("30-Jun-" + varYear);
                        }
                        else if (varQuarter == 1)
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

                        //Get the ending for the QDC extract (number of years since 2003 plus 2 plus the quarter)
                        varQDCID = 1;
                        if (varStartDate.Month == 1)
                            varQDCID = ((varYear - 2003) * 4) + 2 + 1;
                        else if (varStartDate.Month == 4)
                            varQDCID = ((varYear - 2003) * 4) + 2 + 2;
                        else if (varStartDate.Month == 7)
                            varQDCID = ((varYear - 2003) * 4) + 2 + 3;
                        else if (varStartDate.Month == 10)
                            varQDCID = ((varYear - 2003) * 4) + 2 + 4;

                        // Create the file names
                        varFileName = "60600559191_" + varQDCID + ".xml";
                        varFileName2 = "Errors for QDC " + varYear + varQuarter + ".txt";

                        // Fetch statements for database
                        // Get the required fields from the QDC table for the current period
                        string dbQDCList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientqdc'>
                                <attribute name='new_agencyid' />
                                <attribute name='new_contactsfacetoface' />
                                <attribute name='new_contactsother' />
                                <attribute name='new_contactstelephone' />
                                <attribute name='new_description' />
                                <attribute name='new_edtraingroupscarerfamily' />
                                <attribute name='new_edtraingroupscommunity' />
                                <attribute name='new_edtraingroupsconsumer' />
                                <attribute name='new_edtraingroupsprof' />
                                <attribute name='new_edtrainparticipantscarerfamily' />
                                <attribute name='new_edtrainparticipantscommunity' />
                                <attribute name='new_edtrainparticipantsconsumer' />
                                <attribute name='new_edtrainparticipantsprof' />
                                <attribute name='new_msshactivitiesprovided' />
                                <attribute name='new_nbrserviceusers' />
                                <attribute name='new_paidstaffhrsrefweek' />
                                <attribute name='new_palmclientqdcid' />
                                <attribute name='new_participantscarersfamfrnd' />
                                <attribute name='new_participantsconsumers' />
                                <attribute name='new_qualityplan' />
                                <attribute name='new_quarter' />
                                <attribute name='new_sleepovernights' />
                                <attribute name='new_staffengageddirect' />
                                <attribute name='new_staffengagedindirect' />
                                <attribute name='new_supportgroupscarerfamfrnd' />
                                <attribute name='new_supportgroupsconsumers' />
                                <attribute name='new_supportgroupsfemale' />
                                <attribute name='new_unpaidstaffhrsrefweek' />
                                <attribute name='new_waitinglistnbrpeople' />
                                <attribute name='new_year' />
                                <filter type='and'>
                                    <condition entityname='new_palmclientqdc' attribute='new_year' operator='eq' value='" + varYear + @"' />
                                    <condition entityname='new_palmclientqdc' attribute='new_quarter' operator='eq' value='" + varQuarterFull + @"' />
                                </filter>
                                <order attribute='new_description' />
                              </entity>
                            </fetch> ";

                        // Get the required fields from the SU Drop Down list entity
                        string dbDropList = @"
                                <fetch version='1.0' mapping='logical' distinct='false'>
                                  <entity name='new_palmsudropdown'>
                                    <attribute name='new_type' />
                                    <attribute name='new_description' />
                                    <attribute name='new_qdc' />
                                    <order attribute='new_description' />
                                  </entity>
                                </fetch> ";

                        // Get the required fields from the SU QDC Agency entity
                        string dbAgencyList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                                <entity name='new_palmsuqdcagency'>
                                <attribute name='new_agency' />
                                <attribute name='new_agencyid' />
                                <attribute name='new_aihwservicetype' />
                                <attribute name='new_currentcapacity' />
                                <attribute name='new_daysperweek' />
                                <attribute name='new_fullquarteroperation' />
                                <attribute name='new_hoursperday' />
                                <attribute name='new_outletcode' />
                                <attribute name='new_outlettypepdssrr' />
                                <attribute name='new_palmsuqdcagencyid' />
                                <attribute name='new_waitinglistexistance' />
                                <attribute name='new_weeksperquarter' />
                                <order attribute='new_agency' />
                                </entity>
                            </fetch> ";

                        // Get the required fields from the support period (and associated entities)
                        // Any support periods active during the period checked as QDC
                        string dbSupportList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclient'>
                                <attribute name='new_address' />
                                <attribute name='new_palmclientid' />
                                <attribute name='new_firstname' />
                                <attribute name='new_surname' />
                                <attribute name='new_gender' />
                                <attribute name='new_sex' />
                                <attribute name='new_dob' />
                                <attribute name='new_dobest' />
                                <attribute name='new_consent' />
                                <attribute name='new_indigenous' />
                                <attribute name='new_country' />
                                <attribute name='new_language' />
                                <attribute name='new_interpret' />
                                <attribute name='new_communication' />
                                <attribute name='new_depchild' />
                                <link-entity name='new_palmclientsupport' to='new_palmclientid' from='new_client' link-type='inner'>
                                    <attribute name='new_locality' />
                                    <attribute name='new_careravail' />
                                    <attribute name='new_carerprimary' />
                                    <attribute name='new_carerresidence' />
                                    <attribute name='new_carerrship' />
                                    <attribute name='new_carerdob' />
                                    <attribute name='new_carerpdrssclient' />
                                    <attribute name='new_ccpdistype' />
                                    <attribute name='new_otherdis' />
                                    <attribute name='new_primarydiagnosis' />
                                    <attribute name='new_livingarrangepres' />
                                    <attribute name='new_liveswithatexit' />
                                    <attribute name='new_residentialpres' />
                                    <attribute name='new_accommodationexit' />
                                    <attribute name='new_nominationrights' />
                                    <attribute name='new_labourforcepres' />
                                    <attribute name='new_incomepres' />
                                    <attribute name='new_carerallowance' />
                                    <attribute name='new_outsidenotransport' />
                                    <attribute name='new_participatefamily' />
                                    <attribute name='new_participateleisure' />
                                    <attribute name='new_participatemoney' />
                                    <attribute name='new_participatesocial' />
                                    <attribute name='new_participatetransport' />
                                    <attribute name='new_participateworking' />
                                    <attribute name='new_supportcommunication' />
                                    <attribute name='new_supportcommunity' />
                                    <attribute name='new_supportdomestic' />
                                    <attribute name='new_supporteducation' />
                                    <attribute name='new_supportinterpersonal' />
                                    <attribute name='new_supportlearning' />
                                    <attribute name='new_supportmobility' />
                                    <attribute name='new_supportselfcare' />
                                    <attribute name='new_supportworking' />
                                    <attribute name='new_clinicalsupport' />
                                    <attribute name='new_individualfunding' />
                                    <attribute name='new_clienteft' />
                                    <attribute name='new_referraldate' />
                                    <attribute name='new_sourceref' />
                                    <attribute name='new_clientstatus' />
                                    <attribute name='new_cessation' />
                                    <attribute name='new_startdate' />
                                    <attribute name='new_enddate' />
                                    <attribute name='new_doqdc' />
                                    <attribute name='new_palmclientsupportid' />
                                    <link-entity name='new_palmsuqdcagency' to='new_doqdc' from='new_palmsuqdcagencyid' link-type='inner'>
                                        <attribute name='new_outletcode' />
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
                                    <condition entityname='new_palmclientsupport' attribute='new_qdc' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    </filter >
                                </filter>
                                <order attribute='new_address' />
                              </entity>
                            </fetch> ";

                        // Get the required fields from the Activities entity
                        // Any activities marked as QDC against a support period marked as QDC, where the activity occurs in the quarter and the support period is active during the quarter
                        string dbActivityList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                                <entity name='new_palmclientactivities'>
                                    <attribute name='new_supportperiod' />
                                    <attribute name='new_amount' />
                                    <attribute name='new_includeqdc' />
                                    <attribute name='new_entrydate' />
                                    <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                            <attribute name='new_client' />
                                    </link-entity>
                                    <filter type='and'>
                                        <condition entityname='new_palmclientactivities' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                        <condition entityname='new_palmclientactivities' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                        <condition entityname='new_palmclientactivities' attribute='new_includeqdc' operator='eq' value='True' />
                                        <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                        <filter type='or'>
                                            <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                            <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                        </filter >
                                    </filter>
                                </entity>
                            </fetch> ";

                        // Get the required fields from the Case Plans entity
                        // Any case plan against a support period marked as QDC, where the occurs in the quarter and the support period is active during the quarter
                        string dbCasePlanList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                                <entity name='new_palmclientcaseplan'>
                                    <attribute name='new_completed' />
                                    <attribute name='new_createdate' />
                                    <attribute name='new_currstatus' />
                                    <attribute name='new_enddate' />
                                    <attribute name='new_palmclientcaseplanid' />
                                    <attribute name='new_supportperiod' />
                                    <attribute name='new_reviewdate' />
                                    <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                            <attribute name='new_client' />
                                    </link-entity>
                                    <filter type='and'>
                                        <condition entityname='new_palmclientcaseplan' attribute='new_createdate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                        <condition entityname='new_palmclientcaseplan' attribute='new_createdate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                        <condition entityname='new_palmclientsupport' attribute='new_qdc' operator='eq' value='True' />
                                        <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                        <filter type='or'>
                                            <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                            <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                        </filter >
                                    </filter>
                                </entity>
                            </fetch> ";

                        // Variables to hold the data returned from the fetchXML statements
                        string dbAgencyId = "";
                        string dbAgencyId2 = "";
                        string dbContactsFaceToFace = "";
                        string dbContactsOther = "";
                        string dbContactsTelephone = "";
                        string dbDescription = "";
                        string dbEdTrainGroupsCarerFamily = "";
                        string dbEdTrainGroupsCommunity = "";
                        string dbEdTrainGroupsConsumer = "";
                        string dbEdTrainGroupsProf = "";
                        string dbEdTrainParticipantsCarerFamily = "";
                        string dbEdTrainParticipantsCommunity = "";
                        string dbEdTrainParticipantsConsumer = "";
                        string dbEdTrainParticipantsProf = "";
                        string dbMsshActivitiesProvided = "";
                        string dbNbrServiceUsers = "";
                        string dbPaidStaffHrsRefWeek = "";
                        string dbPalmClientQdcId = "";
                        string dbParticipantsCarersFamFrnd = "";
                        string dbParticipantsConsumers = "";
                        string dbQualityPlan = "";
                        string dbQuarter = "";
                        string dbSleepOverNights = "";
                        string dbStaffEngagedDirect = "";
                        string dbStaffEngagedIndirect = "";
                        string dbSupportGroupsCarerFamFrnd = "";
                        string dbSupportGroupsConsumers = "";
                        string dbSupportGroupsFemale = "";
                        string dbUnpaidStaffHrsRefWeek = "";
                        string dbWaitingListNbrPeople = "";
                        string dbYear = "";

                        string dbType = "";
                        string dbQdc = "";

                        string dbCountry = "";
                        string dbLanguage = "";

                        string dbLocality = "";
                        string dbPostcode = "";
                        string dbState = "";

                        string dbAgency = "";
                        string dbAihwServiceType = "";
                        string dbCurrentCapacity = "";
                        string dbDaysPerWeek = "";
                        string dbFullQuarterOperation = "";
                        string dbHoursPerDay = "";
                        string dbOutletCode = "";
                        string dbOutletCode2 = "";
                        string dbOutletTypePdssrr = "";
                        string dbPalmSuQdcAgencyId = "";
                        string dbWaitingListExistance = "";
                        string dbWeeksPerQuarter = "";

                        string dbAddress = "";
                        string dbPalmClientId = "";
                        string dbFirstName = "";
                        string dbSurname = "";
                        string dbGender = "";
                        string dbSex = "";
                        string dbDob = "";
                        string dbDobEst = "";
                        string dbConsent = "";
                        string dbIndigenous = "";
                        string dbInterpret = "";
                        string dbCommunication = "";
                        string dbDepChild = "";

                        string dbCarerAvail = "";
                        string dbCarerPrimary = "";
                        string dbCarerResidence = "";
                        string dbCarerRship = "";
                        string dbCarerDob = "";
                        string dbCarerPdrssClient = "";
                        string dbCCPDisType = "";
                        string dbOtherDis = "";
                        string dbPrimaryDiagnosis = "";
                        string dbLivingArrangePres = "";
                        string dbLivesWithAtExit = "";
                        string dbResidentialPres = "";
                        string dbAccommodationExit = "";
                        string dbNominationRights = "";
                        string dbLabourForcePres = "";
                        string dbIncomePres = "";
                        string dbCarerAllowance = "";
                        string dbOutsideNoTransport = "";
                        string dbParticipateFamily = "";
                        string dbParticipateLeisure = "";
                        string dbParticipateMoney = "";
                        string dbParticipateSocial = "";
                        string dbParticipateTransport = "";
                        string dbParticipateWorking = "";
                        string dbSupportCommunication = "";
                        string dbSupportCommunity = "";
                        string dbSupportDomestic = "";
                        string dbSupportEducation = "";
                        string dbSupportInterpersonal = "";
                        string dbSupportLearning = "";
                        string dbSupportMobility = "";
                        string dbSupportSelfCare = "";
                        string dbSupportWorking = "";
                        string dbClinicalSupport = "";
                        string dbIndividualFunding = "";
                        string dbClientEft = "";
                        string dbReferralDate = "";
                        string dbSourceRef = "";
                        string dbClientStatus = "";
                        string dbCessation = "";
                        string dbStartDate = "";
                        string dbEndDate = "";
                        string dbDoQdc = "";
                        string dbPalmClientSupportId = "";

                        string dbSupportPeriod = "";
                        string dbAmount = "";
                        string dbIncludeQdc = "";
                        string dbEntryDate = "";

                        string dbCompleted = "";
                        string dbCreateDate = "";
                        string dbCurrStatus = "";
                        string dbPalmClientCasePlanId = "";
                        string dbReviewDate = "";
                        string dbCPStartDate = "";

                        string dbClient = "";

                        // Variables for drop down list values
                        string varDesc = "";
                        string varType = "";
                        string varQDC = "";

                        // Avoid duplicate clients
                        int varDoNext = 0;
                        string varClientNumber = "";

                        // Variables for case plan dates
                        string varPPCreateDateAccOutreach = "";
                        string varPPCreateDateCO = "";
                        string varPPCreateDateCongCare = "";
                        string varPPCreateDateDP = "";
                        string varPPCreateDateFamOpt = "";
                        string varPPCreateDateFFYA = "";
                        string varPPCreateDateHomeFirst = "";
                        string varPPCreateDateISP = "";
                        string varPPCreateDateMA = "";
                        string varPPCreateDatePDSS = "";
                        string varPPCreateDateSharedSup = "";
                        string varPPReviewDateAccOutreach = "";
                        string varPPReviewDateCO = "";
                        string varPPReviewDateCongCare = "";
                        string varPPReviewDateDP = "";
                        string varPPReviewDateFamOpt = "";
                        string varPPReviewDate_FFYA = "";
                        string varPPReviewDateHomeFirst = "";
                        string varPPReviewDateISP = "";
                        string varPPReviewDateMA = "";
                        string varPPReviewDatePDSS = "";
                        string varPPReviewDateSharedSup = "";
                        string varPPGoals = "";

                        // Variables for QDC data
                        string varService_Type_Outlet_Codeno = "";
                        string varSurname = "";
                        string varFirstName = "";
                        int varAge = 0;
                        int varCAge = 0;
                        string varDob = "";
                        string varSLK = "";
                        string varPensionBenefitStatus = "";
                        string varContactCaseManger = "";
                        string varReasonRespite = "";
                        string varNightsReceived_PDSSRespite = "";
                        string varNightsReceived_PDSSResiRehab = "";
                        string varLastAssessmentDate = "";
                        string varNbrMealsAtCentre = "";
                        string varNbrMealsAtHome = "";

                        // Service day and snapshot
                        string varLast_Service_Date = "";
                        string varSnapshot_Type = "";
                        DateTime varSnapShotDay = new DateTime();
                        string varDoEnd = "";
                        string varDoCessation = "";
                        string varDoReferral = "";
                        string varDoSnapShot = "";

                        DateTime varM1Start = new DateTime(); //Start date of month 1
                        DateTime varM1End = new DateTime(); //End date of month 1
                        double varM1Hours = 0; //Hours for month 1
                        string varM1First = ""; //First service for month 1
                        string varM1Last = ""; //Last service for month 1
                        DateTime varM2Start = new DateTime(); //Start date of month 2
                        DateTime varM2End = new DateTime(); //End date of month 2
                        double varM2Hours = 0; //Hours for month 2
                        string varM2First = ""; //First service for month 2
                        string varM2Last = ""; //Last service for month 2
                        DateTime varM3Start = new DateTime(); //Start date of month 3
                        DateTime varM3End = new DateTime(); //End date of month 3
                        double varM3Hours = 0; //Hours for month 3
                        string varM3First = ""; //First service for month 3
                        string varM3Last = ""; //Last service for month 3

                        // Get the fetch XML data and place in entity collection objects
                        EntityCollection result = _service.RetrieveMultiple(new FetchExpression(dbQDCList));
                        EntityCollection result2 = _service.RetrieveMultiple(new FetchExpression(dbDropList));
                        EntityCollection result6 = _service.RetrieveMultiple(new FetchExpression(dbAgencyList));
                        EntityCollection result7 = _service.RetrieveMultiple(new FetchExpression(dbSupportList));
                        EntityCollection result8 = _service.RetrieveMultiple(new FetchExpression(dbActivityList));
                        EntityCollection result9 = _service.RetrieveMultiple(new FetchExpression(dbCasePlanList));

                        // Create header of XML file
                        sbHeaderList.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");
                        sbHeaderList.AppendLine("<agency agency_codeno=\"60600\" qdc_id=\"" + varQDCID + "\">");

                        //1/
                        //loop through agencies
                        //This information is in the agency table and edited in the admin page
                        //Set to null or don't do if the type is PDRSS and it is a DSD question etc...
                        foreach (var c in result6.Entities)
                        {
                            // Process the data as follows:
                            // If there is a formatted value for the field, use it
                            // Otherwise if there is a literal value for the field, use it
                            // Otherwise the value wasn't returned so set as nothing
                            if (c.FormattedValues.Contains("new_agencyid"))
                                dbAgencyId = c.FormattedValues["new_agencyid"];
                            else if (c.Attributes.Contains("new_agencyid"))
                                dbAgencyId = c.Attributes["new_agencyid"].ToString();
                            else
                                dbAgencyId = "";

                            if (c.FormattedValues.Contains("new_outletcode"))
                                dbOutletCode = c.FormattedValues["new_outletcode"];
                            else if (c.Attributes.Contains("new_outletcode"))
                                dbOutletCode = c.Attributes["new_outletcode"].ToString();
                            else
                                dbOutletCode = "";

                            if (c.FormattedValues.Contains("new_aihwservicetype"))
                                dbAihwServiceType = c.FormattedValues["new_aihwservicetype"];
                            else if (c.Attributes.Contains("new_aihwservicetype"))
                                dbAihwServiceType = c.Attributes["new_aihwservicetype"].ToString();
                            else
                                dbAihwServiceType = "";

                            if (c.FormattedValues.Contains("new_fullquarteroperation"))
                                dbFullQuarterOperation = c.FormattedValues["new_fullquarteroperation"];
                            else if (c.Attributes.Contains("new_fullquarteroperation"))
                                dbFullQuarterOperation = c.Attributes["new_fullquarteroperation"].ToString();
                            else
                                dbFullQuarterOperation = "";

                            if (c.FormattedValues.Contains("new_weeksperquarter"))
                                dbWeeksPerQuarter = c.FormattedValues["new_weeksperquarter"];
                            else if (c.Attributes.Contains("new_weeksperquarter"))
                                dbWeeksPerQuarter = c.Attributes["new_weeksperquarter"].ToString();
                            else
                                dbWeeksPerQuarter = "";

                            if (c.FormattedValues.Contains("new_daysperweek"))
                                dbDaysPerWeek = c.FormattedValues["new_daysperweek"];
                            else if (c.Attributes.Contains("new_daysperweek"))
                                dbDaysPerWeek = c.Attributes["new_daysperweek"].ToString();
                            else
                                dbDaysPerWeek = "";

                            if (c.FormattedValues.Contains("new_hoursperday"))
                                dbHoursPerDay = c.FormattedValues["new_hoursperday"];
                            else if (c.Attributes.Contains("new_hoursperday"))
                                dbHoursPerDay = c.Attributes["new_hoursperday"].ToString();
                            else
                                dbHoursPerDay = "";

                            if (c.FormattedValues.Contains("new_currentcapacity"))
                                dbCurrentCapacity = c.FormattedValues["new_currentcapacity"];
                            else if (c.Attributes.Contains("new_currentcapacity"))
                                dbCurrentCapacity = c.Attributes["new_currentcapacity"].ToString();
                            else
                                dbCurrentCapacity = "";

                            if (c.FormattedValues.Contains("new_outlettypepdssrr"))
                                dbOutletTypePdssrr = c.FormattedValues["new_outlettypepdssrr"];
                            else if (c.Attributes.Contains("new_outlettypepdssrr"))
                                dbOutletTypePdssrr = c.Attributes["new_outlettypepdssrr"].ToString();
                            else
                                dbOutletTypePdssrr = "";

                            if (c.FormattedValues.Contains("new_waitinglistexistance"))
                                dbWaitingListExistance = c.FormattedValues["new_waitinglistexistance"];
                            else if (c.Attributes.Contains("new_waitinglistexistance"))
                                dbWaitingListExistance = c.Attributes["new_waitinglistexistance"].ToString();
                            else
                                dbWaitingListExistance = "";

                            if (c.FormattedValues.Contains("new_palmsuqdcagencyid"))
                                dbPalmSuQdcAgencyId = c.FormattedValues["new_palmsuqdcagencyid"];
                            else if (c.Attributes.Contains("new_palmsuqdcagencyid"))
                                dbPalmSuQdcAgencyId = c.Attributes["new_palmsuqdcagencyid"].ToString();
                            else
                                dbPalmSuQdcAgencyId = "";

                            //Reset values
                            dbPaidStaffHrsRefWeek = "";
                            dbUnpaidStaffHrsRefWeek = "";
                            dbStaffEngagedDirect = "";
                            dbStaffEngagedIndirect = "";
                            dbSleepOverNights = "";
                            dbNbrServiceUsers = "";
                            dbMsshActivitiesProvided = "";
                            dbContactsFaceToFace = "";
                            dbContactsTelephone = "";
                            dbContactsOther = "";
                            dbSupportGroupsConsumers = "";
                            dbParticipantsConsumers = "";
                            dbSupportGroupsCarerFamFrnd = "";
                            dbParticipantsCarersFamFrnd = "";
                            dbSupportGroupsFemale = "";
                            dbEdTrainGroupsConsumer = "";
                            dbEdTrainParticipantsConsumer = "";
                            dbEdTrainGroupsCarerFamily = "";
                            dbEdTrainParticipantsCarerFamily = "";
                            dbEdTrainGroupsCommunity = "";
                            dbEdTrainParticipantsCommunity = "";
                            dbEdTrainGroupsProf = "";
                            dbEdTrainParticipantsProf = "";
                            dbWaitingListNbrPeople = "";
                            dbQualityPlan = "";

                            //Loop through the agency information and get the correct one for the agency
                            foreach (var q in result.Entities)
                            {
                                // We need to get the entity id for the client field for comparisons
                                if (q.Attributes.Contains("new_agencyid"))
                                {
                                    // Get the entity id for the agency using the entity reference object
                                    getEntity = (EntityReference)q.Attributes["new_agencyid"];
                                    dbAgencyId2 = getEntity.Id.ToString();
                                }
                                else if (q.FormattedValues.Contains("new_agencyid"))
                                    dbAgencyId2 = q.FormattedValues["new_agencyid"];
                                else
                                    dbAgencyId2 = "";

                                // Only get the values if the agencies are the same
                                if (dbAgencyId2 == dbPalmSuQdcAgencyId)
                                {
                                    // Process the data as follows:
                                    // If there is a formatted value for the field, use it
                                    // Otherwise if there is a literal value for the field, use it
                                    // Otherwise the value wasn't returned so set as nothing
                                    if (q.FormattedValues.Contains("new_paidstaffhrsrefweek"))
                                        dbPaidStaffHrsRefWeek = q.FormattedValues["new_paidstaffhrsrefweek"];
                                    else if (q.Attributes.Contains("new_paidstaffhrsrefweek"))
                                        dbPaidStaffHrsRefWeek = q.Attributes["new_paidstaffhrsrefweek"].ToString();
                                    else
                                        dbPaidStaffHrsRefWeek = "";

                                    if (q.FormattedValues.Contains("new_unpaidstaffhrsrefweek"))
                                        dbUnpaidStaffHrsRefWeek = q.FormattedValues["new_unpaidstaffhrsrefweek"];
                                    else if (q.Attributes.Contains("new_unpaidstaffhrsrefweek"))
                                        dbUnpaidStaffHrsRefWeek = q.Attributes["new_unpaidstaffhrsrefweek"].ToString();
                                    else
                                        dbUnpaidStaffHrsRefWeek = "";

                                    if (q.FormattedValues.Contains("new_staffengageddirect"))
                                        dbStaffEngagedDirect = q.FormattedValues["new_staffengageddirect"];
                                    else if (q.Attributes.Contains("new_staffengageddirect"))
                                        dbStaffEngagedDirect = q.Attributes["new_staffengageddirect"].ToString();
                                    else
                                        dbStaffEngagedDirect = "";

                                    if (q.FormattedValues.Contains("new_staffengagedindirect"))
                                        dbStaffEngagedIndirect = q.FormattedValues["new_staffengagedindirect"];
                                    else if (q.Attributes.Contains("new_staffengagedindirect"))
                                        dbStaffEngagedIndirect = q.Attributes["new_staffengagedindirect"].ToString();
                                    else
                                        dbStaffEngagedIndirect = "";

                                    if (q.FormattedValues.Contains("new_sleepovernights"))
                                        dbSleepOverNights = q.FormattedValues["new_sleepovernights"];
                                    else if (q.Attributes.Contains("new_sleepovernights"))
                                        dbSleepOverNights = q.Attributes["new_sleepovernights"].ToString();
                                    else
                                        dbSleepOverNights = "";

                                    if (q.FormattedValues.Contains("new_nbrserviceusers"))
                                        dbNbrServiceUsers = q.FormattedValues["new_nbrserviceusers"];
                                    else if (q.Attributes.Contains("new_nbrserviceusers"))
                                        dbNbrServiceUsers = q.Attributes["new_nbrserviceusers"].ToString();
                                    else
                                        dbNbrServiceUsers = "";

                                    if (q.FormattedValues.Contains("new_msshactivitiesprovided"))
                                        dbMsshActivitiesProvided = q.FormattedValues["new_msshactivitiesprovided"];
                                    else if (q.Attributes.Contains("new_msshactivitiesprovided"))
                                        dbMsshActivitiesProvided = q.Attributes["new_msshactivitiesprovided"].ToString();
                                    else
                                        dbMsshActivitiesProvided = "";

                                    if (q.FormattedValues.Contains("new_contactsfacetoface"))
                                        dbContactsFaceToFace = q.FormattedValues["new_contactsfacetoface"];
                                    else if (q.Attributes.Contains("new_contactsfacetoface"))
                                        dbContactsFaceToFace = q.Attributes["new_contactsfacetoface"].ToString();
                                    else
                                        dbContactsFaceToFace = "";

                                    if (q.FormattedValues.Contains("new_contactstelephone"))
                                        dbContactsTelephone = q.FormattedValues["new_contactstelephone"];
                                    else if (q.Attributes.Contains("new_contactstelephone"))
                                        dbContactsTelephone = q.Attributes["new_contactstelephone"].ToString();
                                    else
                                        dbContactsTelephone = "";

                                    if (q.FormattedValues.Contains("new_contactsother"))
                                        dbContactsOther = q.FormattedValues["new_contactsother"];
                                    else if (q.Attributes.Contains("new_contactsother"))
                                        dbContactsOther = q.Attributes["new_contactsother"].ToString();
                                    else
                                        dbContactsOther = "";

                                    if (q.FormattedValues.Contains("new_supportgroupsconsumers"))
                                        dbSupportGroupsConsumers = q.FormattedValues["new_supportgroupsconsumers"];
                                    else if (q.Attributes.Contains("new_supportgroupsconsumers"))
                                        dbSupportGroupsConsumers = q.Attributes["new_supportgroupsconsumers"].ToString();
                                    else
                                        dbSupportGroupsConsumers = "";

                                    if (q.FormattedValues.Contains("new_participantsconsumers"))
                                        dbParticipantsConsumers = q.FormattedValues["new_participantsconsumers"];
                                    else if (q.Attributes.Contains("new_participantsconsumers"))
                                        dbParticipantsConsumers = q.Attributes["new_participantsconsumers"].ToString();
                                    else
                                        dbParticipantsConsumers = "";

                                    if (q.FormattedValues.Contains("new_supportgroupscarerfamfrnd"))
                                        dbSupportGroupsCarerFamFrnd = q.FormattedValues["new_supportgroupscarerfamfrnd"];
                                    else if (q.Attributes.Contains("new_supportgroupscarerfamfrnd"))
                                        dbSupportGroupsCarerFamFrnd = q.Attributes["new_supportgroupscarerfamfrnd"].ToString();
                                    else
                                        dbSupportGroupsCarerFamFrnd = "";

                                    if (q.FormattedValues.Contains("new_participantscarersfamfrnd"))
                                        dbParticipantsCarersFamFrnd = q.FormattedValues["new_participantscarersfamfrnd"];
                                    else if (q.Attributes.Contains("new_participantscarersfamfrnd"))
                                        dbParticipantsCarersFamFrnd = q.Attributes["new_participantscarersfamfrnd"].ToString();
                                    else
                                        dbParticipantsCarersFamFrnd = "";

                                    if (q.FormattedValues.Contains("new_supportgroupsfemale"))
                                        dbSupportGroupsFemale = q.FormattedValues["new_supportgroupsfemale"];
                                    else if (q.Attributes.Contains("new_supportgroupsfemale"))
                                        dbSupportGroupsFemale = q.Attributes["new_supportgroupsfemale"].ToString();
                                    else
                                        dbSupportGroupsFemale = "";

                                    if (q.FormattedValues.Contains("new_edtraingroupsconsumer"))
                                        dbEdTrainGroupsConsumer = q.FormattedValues["new_edtraingroupsconsumer"];
                                    else if (q.Attributes.Contains("new_edtraingroupsconsumer"))
                                        dbEdTrainGroupsConsumer = q.Attributes["new_edtraingroupsconsumer"].ToString();
                                    else
                                        dbEdTrainGroupsConsumer = "";

                                    if (q.FormattedValues.Contains("new_edtrainparticipantsconsumer"))
                                        dbEdTrainParticipantsConsumer = q.FormattedValues["new_edtrainparticipantsconsumer"];
                                    else if (q.Attributes.Contains("new_edtrainparticipantsconsumer"))
                                        dbEdTrainParticipantsConsumer = q.Attributes["new_edtrainparticipantsconsumer"].ToString();
                                    else
                                        dbEdTrainParticipantsConsumer = "";

                                    if (q.FormattedValues.Contains("new_edtraingroupscarerfamily"))
                                        dbEdTrainGroupsCarerFamily = q.FormattedValues["new_edtraingroupscarerfamily"];
                                    else if (q.Attributes.Contains("new_edtraingroupscarerfamily"))
                                        dbEdTrainGroupsCarerFamily = q.Attributes["new_edtraingroupscarerfamily"].ToString();
                                    else
                                        dbEdTrainGroupsCarerFamily = "";

                                    if (q.FormattedValues.Contains("new_edtrainparticipantscarerfamily"))
                                        dbEdTrainParticipantsCarerFamily = q.FormattedValues["new_edtrainparticipantscarerfamily"];
                                    else if (q.Attributes.Contains("new_edtrainparticipantscarerfamily"))
                                        dbEdTrainParticipantsCarerFamily = q.Attributes["new_edtrainparticipantscarerfamily"].ToString();
                                    else
                                        dbEdTrainParticipantsCarerFamily = "";

                                    if (q.FormattedValues.Contains("new_edtraingroupscommunity"))
                                        dbEdTrainGroupsCommunity = q.FormattedValues["new_edtraingroupscommunity"];
                                    else if (q.Attributes.Contains("new_edtraingroupscommunity"))
                                        dbEdTrainGroupsCommunity = q.Attributes["new_edtraingroupscommunity"].ToString();
                                    else
                                        dbEdTrainGroupsCommunity = "";

                                    if (q.FormattedValues.Contains("new_edtrainparticipantscommunity"))
                                        dbEdTrainParticipantsCommunity = q.FormattedValues["new_edtrainparticipantscommunity"];
                                    else if (q.Attributes.Contains("new_edtrainparticipantscommunity"))
                                        dbEdTrainParticipantsCommunity = q.Attributes["new_edtrainparticipantscommunity"].ToString();
                                    else
                                        dbEdTrainParticipantsCommunity = "";

                                    if (q.FormattedValues.Contains("new_edtraingroupsprof"))
                                        dbEdTrainGroupsProf = q.FormattedValues["new_edtraingroupsprof"];
                                    else if (q.Attributes.Contains("new_edtraingroupsprof"))
                                        dbEdTrainGroupsProf = q.Attributes["new_edtraingroupsprof"].ToString();
                                    else
                                        dbEdTrainGroupsProf = "";

                                    if (q.FormattedValues.Contains("new_edtrainparticipantsprof"))
                                        dbEdTrainParticipantsProf = q.FormattedValues["new_edtrainparticipantsprof"];
                                    else if (q.Attributes.Contains("new_edtrainparticipantsprof"))
                                        dbEdTrainParticipantsProf = q.Attributes["new_edtrainparticipantsprof"].ToString();
                                    else
                                        dbEdTrainParticipantsProf = "";

                                    if (q.FormattedValues.Contains("new_waitinglistnbrpeople"))
                                        dbWaitingListNbrPeople = q.FormattedValues["new_waitinglistnbrpeople"];
                                    else if (q.Attributes.Contains("new_waitinglistnbrpeople"))
                                        dbWaitingListNbrPeople = q.Attributes["new_waitinglistnbrpeople"].ToString();
                                    else
                                        dbWaitingListNbrPeople = "";

                                    if (q.FormattedValues.Contains("new_qualityplan"))
                                        dbQualityPlan = q.FormattedValues["new_qualityplan"];
                                    else if (q.Attributes.Contains("new_qualityplan"))
                                        dbQualityPlan = q.Attributes["new_qualityplan"].ToString();
                                    else
                                        dbQualityPlan = "";

                                    break;
                                } // Same agency

                            } // QDC Loop

                            //Get the agency value based on the QDC code
                            varAgencyValue = 0;

                            if (String.IsNullOrEmpty(dbOutletCode) == true)
                                dbOutletCode = "00";

                            if (dbOutletCode.Length > 1)
                            {
                                if (dbOutletCode.Substring(0, 2) == "17") //DSD
                                    varAgencyValue = 1;
                                else if (dbOutletCode.Substring(0, 2) == "15") //PDRSS
                                    varAgencyValue = 2;
                                else if (dbOutletCode.Substring(0, 2) == "13") //HACC
                                    varAgencyValue = 3;
                                else
                                {
                                    varAgencyValue = 0;
                                    sbErrorList.AppendLine("Error: Agency type not defined - " + dbAgencyId + "<br>");
                                }
                            }
                            else
                            {
                                varAgencyValue = 0;
                                sbErrorList.AppendLine("Error: Agency type not defined - " + dbAgencyId + "<br>");
                            }

                            // Loop through drop down list data
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

                                if (d.FormattedValues.Contains("new_qdc"))
                                    varQDC = d.FormattedValues["new_qdc"];
                                else if (d.Attributes.Contains("new_qdc"))
                                    varQDC = d.Attributes["new_qdc"].ToString();
                                else
                                    varQDC = "";

                                // Ensure numeric
                                varQDC = cleanString(varQDC, "number");
                                if (String.IsNullOrEmpty(varQDC) == true)
                                    varQDC = "0";

                                // If the type is servicetype and the values match, get the QDC value
                                if (varType.ToLower() == "servicetype")
                                {
                                    if (dbAihwServiceType.ToLower() == varDesc.ToLower())
                                        dbAihwServiceType = varQDC;
                                }
                            } // Drop list

                            //Error if agency id is null
                            if (String.IsNullOrEmpty(dbAgencyId) == true)
                            {
                                dbAgencyId = "0";
                                sbErrorList.AppendLine("Error: Agency Outlet Code not defined - " + dbAgencyId + "<br>");
                            }

                            //For DSD and PDRSS only
                            if (varAgencyValue == 1 || varAgencyValue == 2)
                            {
                                //Error if no service type
                                if (String.IsNullOrEmpty(dbAihwServiceType) == true)
                                    sbErrorList.AppendLine("Error: AIHW service type not defined - " + dbAgencyId + "<br>");

                                //Convert to number or alert of error if null
                                if (dbFullQuarterOperation == "Yes")
                                    dbFullQuarterOperation = "1";
                                else if (dbFullQuarterOperation == "No")
                                    dbFullQuarterOperation = "2";
                                else
                                    sbErrorList.AppendLine("Error: Full quarter operation not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 13
                                if (String.IsNullOrEmpty(dbWeeksPerQuarter) == false)
                                {
                                    if (Convert.ToDouble(dbWeeksPerQuarter) < 1 || (Convert.ToDouble(dbWeeksPerQuarter) > 13 && Convert.ToDouble(dbWeeksPerQuarter) != 90))
                                        sbErrorList.AppendLine("Error: Weeks per quarter has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Weeks per quarter not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 7
                                if (String.IsNullOrEmpty(dbDaysPerWeek) == false)
                                {
                                    if (Convert.ToDouble(dbDaysPerWeek) < 1 || (Convert.ToDouble(dbDaysPerWeek) > 7 && Convert.ToDouble(dbDaysPerWeek) != 90))
                                        sbErrorList.AppendLine("Error: Days per week has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Days per week not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 24
                                if (String.IsNullOrEmpty(dbHoursPerDay) == false)
                                {
                                    if (Convert.ToDouble(dbHoursPerDay) < 1 || (Convert.ToDouble(dbHoursPerDay) > 24 && Convert.ToDouble(dbHoursPerDay) != 90))
                                        sbErrorList.AppendLine("Error: Hours per day has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Hours per day not defined - " + dbAgencyId + "<br>");

                            } //agencyvalue 1 or 2


                            //For specific agency types only
                            if (dbOutletCode == "17010" || dbOutletCode == "17016" || dbOutletCode == "15037" || dbOutletCode == "15038" || dbOutletCode == "15055")
                            {

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbCurrentCapacity) == false)
                                {
                                    if (Convert.ToDouble(dbCurrentCapacity) < 1 || Convert.ToDouble(dbCurrentCapacity) > 999)
                                        sbErrorList.AppendLine("Error: Current capacity has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Current capacity not defined - " + dbAgencyId + "<br>");

                            }


                            //For specific agency type only
                            if (dbOutletCode == "15038")
                            {

                                //Get a number based on the value or alert of error
                                if (dbOutletTypePdssrr == "Young Persons Residential rehabilitation Program")
                                    dbOutletTypePdssrr = "1";
                                else if (dbOutletTypePdssrr == "Long Term Residential rehabilitation Program")
                                    dbOutletTypePdssrr = "2";
                                else
                                    sbErrorList.AppendLine("Error: Residential rehabilitation program not defined - " + dbAgencyId + "<br>");

                            }


                            //For DSD and PDRSS only
                            if (varAgencyValue == 1 || varAgencyValue == 2)
                            {

                                //Check for valid number between 1 and 9999
                                if (String.IsNullOrEmpty(dbPaidStaffHrsRefWeek) == false)
                                {
                                    if (Convert.ToDouble(dbPaidStaffHrsRefWeek) < 0 || Convert.ToDouble(dbPaidStaffHrsRefWeek) > 9999)
                                        sbErrorList.AppendLine("Error: Paid staff hours has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Paid staff hours not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 9999
                                if (String.IsNullOrEmpty(dbUnpaidStaffHrsRefWeek) == false)
                                {
                                    if (Convert.ToDouble(dbUnpaidStaffHrsRefWeek) < 0 || Convert.ToDouble(dbUnpaidStaffHrsRefWeek) > 9999)
                                        sbErrorList.AppendLine("Error: Unpaid staff hours has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Unpaid staff hours not defined - " + dbAgencyId + "<br>");

                            }


                            //For PDRSS only
                            if (varAgencyValue == 2)
                            {

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbStaffEngagedDirect) == false)
                                {
                                    if (Convert.ToDouble(dbStaffEngagedDirect) < 0 || Convert.ToDouble(dbStaffEngagedDirect) > 999)
                                        sbErrorList.AppendLine("Error: Direct staff hours has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Direct staff hours not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbStaffEngagedIndirect) == false)
                                {
                                    if (Convert.ToDouble(dbStaffEngagedIndirect) < 0 || Convert.ToDouble(dbStaffEngagedIndirect) > 999)
                                        sbErrorList.AppendLine("Error: Indirect staff hours has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Indirect staff hours not defined - " + dbAgencyId + "<br>");

                            }


                            //For specific agency types only
                            if (dbOutletCode == "15037" || dbOutletCode == "15038" || dbOutletCode == "15055")
                            {

                                //Check for valid number between 0 and 99
                                if (String.IsNullOrEmpty(dbSleepOverNights) == false)
                                {
                                    if (Convert.ToDouble(dbSleepOverNights) < 0 || Convert.ToDouble(dbSleepOverNights) > 99)
                                        sbErrorList.AppendLine("Error: Staff working overnight has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Staff working overnight not defined - " + dbAgencyId + "<br>");

                            }


                            //For specific agency types only
                            if (dbOutletCode == "17024" || dbOutletCode == "17033" || dbOutletCode == "17044" || dbOutletCode == "15263" || dbOutletCode == "15036")
                            {

                                //Check for valid number between 0 and 9999
                                if (String.IsNullOrEmpty(dbNbrServiceUsers) == false)
                                {
                                    if (Convert.ToDouble(dbNbrServiceUsers) < 0 || Convert.ToDouble(dbNbrServiceUsers) > 9999)
                                        sbErrorList.AppendLine("Error: Number of service users has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Number of service users not defined - " + dbAgencyId + "<br>");

                            }


                            //For specific agency types only
                            if (dbOutletCode == "15263" || dbOutletCode == "15036")
                            {

                                //Get number based on value or alert of error
                                if (dbMsshActivitiesProvided == "Information/Advice")
                                    dbMsshActivitiesProvided = "1";
                                else if (dbMsshActivitiesProvided == "Telephone")
                                    dbMsshActivitiesProvided = "2";
                                else if (dbMsshActivitiesProvided == "Support groups")
                                    dbMsshActivitiesProvided = "3";
                                else if (dbMsshActivitiesProvided == "Education forums")
                                    dbMsshActivitiesProvided = "4";
                                else
                                    sbErrorList.AppendLine("Error: Mutual support and self help activities not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbContactsFaceToFace) == false)
                                {
                                    if (Convert.ToDouble(dbContactsFaceToFace) < 1 || Convert.ToDouble(dbContactsFaceToFace) > 999)
                                        sbErrorList.AppendLine("Error: Face to face contacts has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Face to face contacts not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbContactsTelephone) == false)
                                {
                                    if (Convert.ToDouble(dbContactsTelephone) < 1 || Convert.ToDouble(dbContactsTelephone) > 999)
                                        sbErrorList.AppendLine("Error: Telephone contacts has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Telephone contacts not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbContactsOther) == false)
                                {
                                    if (Convert.ToDouble(dbContactsOther) < 1 || Convert.ToDouble(dbContactsOther) > 999)
                                        sbErrorList.AppendLine("Error: Other contacts has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Other contacts not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbSupportGroupsConsumers) == false)
                                {
                                    if (Convert.ToDouble(dbSupportGroupsConsumers) < 1 || Convert.ToDouble(dbSupportGroupsConsumers) > 999)
                                        sbErrorList.AppendLine("Error: Number of consumer support groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Number of consumer support groups not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbParticipantsConsumers) == false)
                                {
                                    if (Convert.ToDouble(dbParticipantsConsumers) < 1 || Convert.ToDouble(dbParticipantsConsumers) > 999)
                                        sbErrorList.AppendLine("Error: Attendee number of consumer support groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Attendee number of consumer support groups not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbSupportGroupsCarerFamFrnd) == false)
                                {
                                    if (Convert.ToDouble(dbSupportGroupsCarerFamFrnd) < 1 || Convert.ToDouble(dbSupportGroupsCarerFamFrnd) > 999)
                                        sbErrorList.AppendLine("Error: Number of carer/family/friend groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Number of carer/family/friend groups not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbParticipantsCarersFamFrnd) == false)
                                {
                                    if (Convert.ToDouble(dbParticipantsCarersFamFrnd) < 1 || Convert.ToDouble(dbParticipantsCarersFamFrnd) > 999)
                                        sbErrorList.AppendLine("Error: Attendee number of carer/family/friend groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Attendee number of carer/family/friend groups not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbSupportGroupsFemale) == false)
                                {
                                    if (Convert.ToDouble(dbSupportGroupsFemale) < 1 || Convert.ToDouble(dbSupportGroupsFemale) > 999)
                                        sbErrorList.AppendLine("Error: Percent of female support group participants has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Percent of female support group participants not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbEdTrainGroupsConsumer) == false)
                                {
                                    if (Convert.ToDouble(dbEdTrainGroupsConsumer) < 1 || Convert.ToDouble(dbEdTrainGroupsConsumer) > 999)
                                        sbErrorList.AppendLine("Error: Number of education or training groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Number of education or training groups not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbEdTrainParticipantsConsumer) == false)
                                {
                                    if (Convert.ToDouble(dbEdTrainParticipantsConsumer) < 1 || Convert.ToDouble(dbEdTrainParticipantsConsumer) > 999)
                                        sbErrorList.AppendLine("Error: Attendee number of education or training groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Attendee number of education or training groups not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbEdTrainGroupsCarerFamily) == false)
                                {
                                    if (Convert.ToDouble(dbEdTrainGroupsCarerFamily) < 1 || Convert.ToDouble(dbEdTrainGroupsCarerFamily) > 999)
                                        sbErrorList.AppendLine("Error: Number of carer education or training groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Number of carer education or training groups not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbEdTrainParticipantsCarerFamily) == false)
                                {
                                    if (Convert.ToDouble(dbEdTrainParticipantsCarerFamily) < 1 || Convert.ToDouble(dbEdTrainParticipantsCarerFamily) > 999)
                                        sbErrorList.AppendLine("Error: Attendee number of carer education or training groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Attendee number of carer education or training groups not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbEdTrainGroupsCommunity) == false)
                                {
                                    if (Convert.ToDouble(dbEdTrainGroupsCommunity) < 1 || Convert.ToDouble(dbEdTrainGroupsCommunity) > 999)
                                        sbErrorList.AppendLine("Error: Number community education or training groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Number community education or training groups not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbEdTrainParticipantsCommunity) == false)
                                {
                                    if (Convert.ToDouble(dbEdTrainParticipantsCommunity) < 1 || Convert.ToDouble(dbEdTrainParticipantsCommunity) > 999)
                                        sbErrorList.AppendLine("Error: Attendee number of community education or training groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Attendee number of community education or training groups not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbEdTrainGroupsProf) == false)
                                {
                                    if (Convert.ToDouble(dbEdTrainGroupsProf) < 1 || Convert.ToDouble(dbEdTrainGroupsProf) > 999)
                                        sbErrorList.AppendLine("Error: Number of professional education or training groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Number of professional education or training groups not defined - " + dbAgencyId + "<br>");

                                //Check for valid number between 1 and 999
                                if (String.IsNullOrEmpty(dbEdTrainParticipantsProf) == false)
                                {
                                    if (Convert.ToDouble(dbEdTrainParticipantsProf) < 1 || Convert.ToDouble(dbEdTrainParticipantsProf) > 999)
                                        sbErrorList.AppendLine("Error: Attendee number of professional education or training groups has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Attendee number of professional education or training groups not defined - " + dbAgencyId + "<br>");

                            } //Specific agency types only


                            //For PDRSS Only
                            if (varAgencyValue == 2)
                            {

                                //Get number based on value or alert of error
                                if (dbWaitingListExistance == "Yes")
                                    dbWaitingListExistance = "1";
                                else if (dbWaitingListExistance == "No")
                                    dbWaitingListExistance = "2";
                                else
                                    sbErrorList.AppendLine("Error: Existance of waiting list not defined - " + dbAgencyId + "<br>");

                            }


                            //For specific agency type or PDRSS only
                            if (dbOutletCode == "17028" || varAgencyValue == 2)
                            {

                                //Check for valid number between 0 and 99
                                if (String.IsNullOrEmpty(dbWaitingListNbrPeople) == false)
                                {
                                    if (Convert.ToDouble(dbWaitingListNbrPeople) < 0 || Convert.ToDouble(dbWaitingListNbrPeople) > 99)
                                        sbErrorList.AppendLine("Error: Number of people on waiting list has an invalid value - " + dbAgencyId + "<br>");
                                }
                                else
                                    sbErrorList.AppendLine("Error: Number of people on waiting list not defined - " + dbAgencyId + "<br>");

                            }


                            //For DSD only
                            if (varAgencyValue == 1)
                            {
                                //Alert of error if not valid number
                                if (dbQualityPlan != "1" && dbQualityPlan != "2" && dbQualityPlan != "3" && dbQualityPlan != "4")
                                    sbErrorList.AppendLine("Error: Quality Plan not defined - " + dbAgencyId + "<br>");

                            }



                            //Create first part of extract
                            //Only insert a line if it is specific for this agency (as defined above)

                            //DSD and PDRSS
                            if (varAgencyValue == 1 || varAgencyValue == 2)
                            {
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"AIHWServiceType\" value=\"" + dbAihwServiceType + "\" />"); //[QDC S01]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"FullQuarterOperation\" value=\"" + dbFullQuarterOperation + "\" />"); //[QDC S02]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"WeeksPerQuarter\" value=\"" + dbWeeksPerQuarter + "\" />"); //[QDC S03]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"DaysPerWeek\" value=\"" + dbDaysPerWeek + "\" />"); //[QDC S04]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"HoursPerDay\" value=\"" + dbHoursPerDay + "\" />"); //[QDC S05]
                            }

                            //Specific type only
                            if (dbOutletCode == "17010" || dbOutletCode == "17016" || dbOutletCode == "15037" || dbOutletCode == "15038" || dbOutletCode == "15055")
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"CurrentCapacity\" value=\"" + dbCurrentCapacity + "\" />"); //[QDC S06]

                            //Specific type only
                            if (dbOutletCode == "15038")
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"OutletTypePDSS-RR\" value=\"" + dbOutletTypePdssrr + "\" />"); //[QDC S07]

                            //DSD and PDRSS
                            if (varAgencyValue == 1 || varAgencyValue == 2)
                            {
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"PaidStaffHrsRefWeek\" value=\"" + dbPaidStaffHrsRefWeek + "\" />"); //[QDC S08]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"UnPaidStaffHrsRefWeek\" value=\"" + dbUnpaidStaffHrsRefWeek + "\" />"); //[QDC S09]
                            }

                            //PDRSS Only
                            if (varAgencyValue == 2)
                            {
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"StaffEngagedDirect\" value=\"" + dbStaffEngagedDirect + "\" />"); //[QDC S10]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"StaffEngagedIndirect\" value=\"" + dbStaffEngagedIndirect + "\" />"); //[QDC S11]
                            }

                            //Specific type only
                            if (dbOutletCode == "15037" || dbOutletCode == "15038" || dbOutletCode == "15055")
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"SleepOverNights\" value=\"" + dbSleepOverNights + "\" />"); //[QDC S12]

                            //DSD and PDRSS
                            if (dbOutletCode == "17024" || dbOutletCode == "17033" || dbOutletCode == "17044" || dbOutletCode == "15263" || dbOutletCode == "15036")
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"NbrServiceUsers\" value=\"" + dbNbrServiceUsers + "\" />"); //[QDC S13]

                            //Specific type only
                            if (dbOutletCode == "15263" || dbOutletCode == "15036")
                            {
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"MSSHActivitiesProvided\" value=\"" + dbMsshActivitiesProvided + "\" />"); //[QDC S14]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"ContactsFaceToFace\" value=\"" + dbContactsFaceToFace + "\" />"); //[QDC S15]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"ContactsTelephone\" value=\"" + dbContactsTelephone + "\" />"); //[QDC S16]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"ContactsOther\" value=\"" + dbContactsOther + "\" />"); //[QDC S17]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"SupportGroupsConsumers\" value=\"" + dbSupportGroupsConsumers + "\" />"); //[QDC S18]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"ParticipantsConsumers\" value=\"" + dbParticipantsConsumers + "\" />"); //[QDC S19]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"SupportGroupsCarer-Fam-Frnd\" value=\"" + dbSupportGroupsCarerFamFrnd + "\" />"); //[QDC S20]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"ParticipantsCarers-Fam-Frnd\" value=\"" + dbParticipantsCarersFamFrnd + "\" />"); //[QDC S21]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"SupportGroupsFemale\" value=\"" + dbSupportGroupsFemale + "\" />"); //[QDC S22]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"EdTrainGroupsConsumer\" value=\"" + dbEdTrainGroupsConsumer + "\" />"); //[QDC S23]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"EdTrainParticipantsConsumer\" value=\"" + dbEdTrainParticipantsConsumer + "\" />"); //[QDC S24]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"EdTrainGroupsCarerFamily\" value=\"" + dbEdTrainGroupsCarerFamily + "\" />"); //[QDC S25]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"EdTrainParticipantsCarerFamily\" value=\"" + dbEdTrainParticipantsCarerFamily + "\" />"); //[QDC S26]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"EdTrainGroupsCommunity\" value=\"" + dbEdTrainGroupsCommunity + "\" />"); //[QDC S27]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"EdTrainParticipantsCommunity\" value=\"" + dbEdTrainParticipantsCommunity + "\" />"); //[QDC S28]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"EdTrainGroupsProf\" value=\"" + dbEdTrainGroupsProf + "\" />"); //[QDC S29]
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"EdTrainParticipantsProf\" value=\"" + dbEdTrainParticipantsProf + "\" />"); //[QDC S30]
                            }

                            //PDRSS Only
                            if (varAgencyValue == 2)
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"WaitingListExistence\" value=\"" + dbWaitingListExistance + "\" />"); //[QDC S31]

                            //Specific Type Only
                            if (dbOutletCode == "17028" || varAgencyValue == 2)
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"WaitingListNbrPeople\" value=\"" + dbWaitingListNbrPeople + "\" />"); //[QDC S32]

                            //DSD Only
                            if (varAgencyValue == 1)
                                sbHeaderList.AppendLine("   <outlet_response service_type_outlet_codeno=\"" + dbAgencyId + "\" question_fieldname=\"QualityPlan\" value=\"" + dbQualityPlan + "\" />"); //[QDC ???]

                        } // Agency Loop

                        // Reset client number duplicate checker
                        varClientNumber = "...";

                        //2/
                        //For each agency get the clients
                        //This data comes from the client details and support period
                        foreach (var c in result7.Entities)
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

                            //Reset duplicate client test
                            varDoNext = 0;
                            if (varClientNumber.IndexOf("*" + dbPalmClientId + "*") > -1)
                                varDoNext = 1;
                            // Add client number to duplicate checker
                            varClientNumber += "*" + dbPalmClientId + "*";

                            //This is a new client - proceed
                            if (varDoNext == 0)
                            {
                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_palmclientsupportid"))
                                    dbPalmClientSupportId = c.FormattedValues["new_palmclientsupport1.new_palmclientsupportid"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_palmclientsupportid"))
                                    dbPalmClientSupportId = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_palmclientsupportid").Value.ToString();
                                else
                                    dbPalmClientSupportId = "";

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

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_careravail"))
                                    dbCarerAvail = c.FormattedValues["new_palmclientsupport1.new_careravail"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_careravail"))
                                    dbCarerAvail = c.Attributes["new_palmclientsupport1.new_careravail"].ToString();
                                else
                                    dbCarerAvail = "";

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

                                if (c.FormattedValues.Contains("new_consent"))
                                    dbConsent = c.FormattedValues["new_consent"];
                                else if (c.Attributes.Contains("new_consent"))
                                    dbConsent = c.Attributes["new_consent"].ToString();
                                else
                                    dbConsent = "";

                                //At some point have an equipment table with issue date and code.
                                //arrIssue_Date [Not sure]
                                //arrEquipment_Code [Not sure]

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_carerprimary"))
                                    dbCarerPrimary = c.FormattedValues["new_palmclientsupport1.new_carerprimary"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_carerprimary"))
                                    dbCarerPrimary = c.Attributes["new_palmclientsupport1.new_carerprimary"].ToString();
                                else
                                    dbCarerPrimary = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_carerresidence"))
                                    dbCarerResidence = c.FormattedValues["new_palmclientsupport1.new_carerresidence"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_carerresidence"))
                                    dbCarerResidence = c.Attributes["new_palmclientsupport1.new_carerresidence"].ToString();
                                else
                                    dbCarerResidence = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_carerrship"))
                                    dbCarerRship = c.FormattedValues["new_palmclientsupport1.new_carerrship"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_carerrship"))
                                    dbCarerRship = c.Attributes["new_palmclientsupport1.new_carerrship"].ToString();
                                else
                                    dbCarerRship = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_carerdob"))
                                    dbCarerDob = c.FormattedValues["new_palmclientsupport1.new_carerdob"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_caredob"))
                                    dbCarerDob = c.Attributes["new_palmclientsupport1.new_carerdob"].ToString();
                                else
                                    dbCarerDob = "";

                                // Convert date from American format to Australian format
                                dbCarerDob = cleanDateAM(dbCarerDob);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_carerpdrssclient"))
                                    dbCarerPdrssClient = c.FormattedValues["new_palmclientsupport1.new_carerpdrssclient"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_carerpdrssclient"))
                                    dbCarerPdrssClient = c.Attributes["new_palmclientsupport1.new_carerpdrssclient"].ToString();
                                else
                                    dbCarerPdrssClient = "";

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

                                if (c.FormattedValues.Contains("new_palmddllanguage5.new_code"))
                                    dbLanguage = c.FormattedValues["new_palmddllanguage5.new_code"];
                                else if (c.Attributes.Contains("new_palmddllanguage5.new_code"))
                                    dbLanguage = c.Attributes["new_palmddllanguage5.new_code"].ToString();
                                else
                                    dbLanguage = "";

                                if (c.FormattedValues.Contains("new_interpret"))
                                    dbInterpret = c.FormattedValues["new_interpret"];
                                else if (c.Attributes.Contains("new_interpret"))
                                    dbInterpret = c.Attributes["new_interpret"].ToString();
                                else
                                    dbInterpret = "";

                                if (c.FormattedValues.Contains("new_communication"))
                                    dbCommunication = c.FormattedValues["new_communication"];
                                else if (c.Attributes.Contains("new_communication"))
                                    dbCommunication = c.Attributes["new_communication"].ToString();
                                else
                                    dbCommunication = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_ccpdistype"))
                                    dbCCPDisType = c.FormattedValues["new_palmclientsupport1.new_ccpdistype"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_ccpdistype"))
                                    dbCCPDisType = c.Attributes["new_palmclientsupport1.new_ccpdistype"].ToString();
                                else
                                    dbCCPDisType = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_otherdis"))
                                    dbOtherDis = c.FormattedValues["new_palmclientsupport1.new_otherdis"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_otherdis"))
                                    dbOtherDis = c.Attributes["new_palmclientsupport1.new_otherdis"].ToString();
                                else
                                    dbOtherDis = "";

                                // Wrap asterisks around values for better comparison
                                dbOtherDis = getMult(dbOtherDis);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_primarydiagnosis"))
                                    dbPrimaryDiagnosis = c.FormattedValues["new_palmclientsupport1.new_primarydiagnosis"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_primarydiagnosis"))
                                    dbPrimaryDiagnosis = c.Attributes["new_palmclientsupport1.new_primarydiagnosis"].ToString();
                                else
                                    dbPrimaryDiagnosis = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_livingarrangepres"))
                                    dbLivingArrangePres = c.FormattedValues["new_palmclientsupport1.new_livingarrangepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_livingarrangepres"))
                                    dbLivingArrangePres = c.Attributes["new_palmclientsupport1.new_livingarrangepres"].ToString();
                                else
                                    dbLivingArrangePres = "";

                                if (c.FormattedValues.Contains("new_depchild"))
                                    dbDepChild = c.FormattedValues["new_depchild"];
                                else if (c.Attributes.Contains("new_depchild"))
                                    dbDepChild = c.Attributes["new_depchild"].ToString();
                                else
                                    dbDepChild = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_liveswithatexit"))
                                    dbLivesWithAtExit = c.FormattedValues["new_palmclientsupport1.new_liveswithatexit"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_liveswithatexit"))
                                    dbLivesWithAtExit = c.Attributes["new_palmclientsupport1.new_liveswithatexit"].ToString();
                                else
                                    dbLivesWithAtExit = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_residentialpres"))
                                    dbResidentialPres = c.FormattedValues["new_palmclientsupport1.new_residentialpres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_residentialpres"))
                                    dbResidentialPres = c.Attributes["new_palmclientsupport1.new_residentialpres"].ToString();
                                else
                                    dbResidentialPres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_accommodationexit"))
                                    dbAccommodationExit = c.FormattedValues["new_palmclientsupport1.new_accommodationexit"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_accommodationexit"))
                                    dbAccommodationExit = c.Attributes["new_palmclientsupport1.new_accommodationexit"].ToString();
                                else
                                    dbAccommodationExit = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_nominationrights"))
                                    dbNominationRights = c.FormattedValues["new_palmclientsupport1.new_nominationrights"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_nominationrights"))
                                    dbNominationRights = c.Attributes["new_palmclientsupport1.new_nominationrights"].ToString();
                                else
                                    dbNominationRights = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_labourforcepres"))
                                    dbLabourForcePres = c.FormattedValues["new_palmclientsupport1.new_labourforcepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_labourforcepres"))
                                    dbLabourForcePres = c.Attributes["new_palmclientsupport1.new_labourforcepres"].ToString();
                                else
                                    dbLabourForcePres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_incomepres"))
                                    dbIncomePres = c.FormattedValues["new_palmclientsupport1.new_incomepres"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_incomepres"))
                                    dbIncomePres = c.Attributes["new_palmclientsupport1.new_incomepres"].ToString();
                                else
                                    dbIncomePres = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_carerallowance"))
                                    dbCarerAllowance = c.FormattedValues["new_palmclientsupport1.new_carerallowance"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_carerallowance"))
                                    dbCarerAllowance = c.Attributes["new_palmclientsupport1.new_carerallowance"].ToString();
                                else
                                    dbCarerAllowance = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_outsidenotransport"))
                                    dbOutsideNoTransport = c.FormattedValues["new_palmclientsupport1.new_outsidenotransport"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_outsidenotransport"))
                                    dbOutsideNoTransport = c.Attributes["new_palmclientsupport1.new_outsidenotransport"].ToString();
                                else
                                    dbOutsideNoTransport = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_participatetransport"))
                                    dbParticipateTransport = c.FormattedValues["new_palmclientsupport1.new_participatetransport"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_participatetransport"))
                                    dbParticipateTransport = c.Attributes["new_palmclientsupport1.new_participatetransport"].ToString();
                                else
                                    dbParticipateTransport = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_participatefamily"))
                                    dbParticipateFamily = c.FormattedValues["new_palmclientsupport1.new_participatefamily"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_participatefamily"))
                                    dbParticipateFamily = c.Attributes["new_palmclientsupport1.new_participatefamily"].ToString();
                                else
                                    dbParticipateFamily = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_participatesocial"))
                                    dbParticipateSocial = c.FormattedValues["new_palmclientsupport1.new_participatesocial"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_participatesocial"))
                                    dbParticipateSocial = c.Attributes["new_palmclientsupport1.new_participatesocial"].ToString();
                                else
                                    dbParticipateSocial = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_participateleisure"))
                                    dbParticipateLeisure = c.FormattedValues["new_palmclientsupport1.new_participateleisure"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_participateleisure"))
                                    dbParticipateLeisure = c.Attributes["new_palmclientsupport1.new_participateleisure"].ToString();
                                else
                                    dbParticipateLeisure = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_participateworking"))
                                    dbParticipateWorking = c.FormattedValues["new_palmclientsupport1.new_participateworking"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_participateworking"))
                                    dbParticipateWorking = c.Attributes["new_palmclientsupport1.new_participateworking"].ToString();
                                else
                                    dbParticipateWorking = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_participatemoney"))
                                    dbParticipateMoney = c.FormattedValues["new_palmclientsupport1.new_participatemoney"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_participatemoney"))
                                    dbParticipateMoney = c.Attributes["new_palmclientsupport1.new_participatemoney"].ToString();
                                else
                                    dbParticipateMoney = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supportselfcare"))
                                    dbSupportSelfCare = c.FormattedValues["new_palmclientsupport1.new_supportselfcare"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supportselfcare"))
                                    dbSupportSelfCare = c.Attributes["new_palmclientsupport1.new_supportselfcare"].ToString();
                                else
                                    dbSupportSelfCare = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supportmobility"))
                                    dbSupportMobility = c.FormattedValues["new_palmclientsupport1.new_supportmobility"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supportmobility"))
                                    dbSupportMobility = c.Attributes["new_palmclientsupport1.new_supportmobility"].ToString();
                                else
                                    dbSupportMobility = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supportcommunication"))
                                    dbSupportCommunication = c.FormattedValues["new_palmclientsupport1.new_supportcommunication"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supportcommunication"))
                                    dbSupportCommunication = c.Attributes["new_palmclientsupport1.new_supportcommunication"].ToString();
                                else
                                    dbSupportCommunication = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supportinterpersonal"))
                                    dbSupportInterpersonal = c.FormattedValues["new_palmclientsupport1.new_supportinterpersonal"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supportinterpersonal"))
                                    dbSupportInterpersonal = c.Attributes["new_palmclientsupport1.new_supportinterpersonal"].ToString();
                                else
                                    dbSupportInterpersonal = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supportlearning"))
                                    dbSupportLearning = c.FormattedValues["new_palmclientsupport1.new_supportlearning"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supportlearning"))
                                    dbSupportLearning = c.Attributes["new_palmclientsupport1.new_supportlearning"].ToString();
                                else
                                    dbSupportLearning = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supporteducation"))
                                    dbSupportEducation = c.FormattedValues["new_palmclientsupport1.new_supporteducation"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supporteducation"))
                                    dbSupportEducation = c.Attributes["new_palmclientsupport1.new_supporteducation"].ToString();
                                else
                                    dbSupportEducation = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supportcommunity"))
                                    dbSupportCommunity = c.FormattedValues["new_palmclientsupport1.new_supportcommunity"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supportcommunity"))
                                    dbSupportCommunity = c.Attributes["new_palmclientsupport1.new_supportcommunity"].ToString();
                                else
                                    dbSupportCommunity = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supportdomestic"))
                                    dbSupportDomestic = c.FormattedValues["new_palmclientsupport1.new_supportdomestic"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supportdomestic"))
                                    dbSupportDomestic = c.Attributes["new_palmclientsupport1.new_supportdomestic"].ToString();
                                else
                                    dbSupportDomestic = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supportworking"))
                                    dbSupportWorking = c.FormattedValues["new_palmclientsupport1.new_supportworking"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supportworking"))
                                    dbSupportWorking = c.Attributes["new_palmclientsupport1.new_supportworking"].ToString();
                                else
                                    dbSupportWorking = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_clinicalsupport"))
                                    dbClinicalSupport = c.FormattedValues["new_palmclientsupport1.new_clinicalsupport"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_clinicalsupport"))
                                    dbClinicalSupport = c.Attributes["new_palmclientsupport1.new_clinicalsupport"].ToString();
                                else
                                    dbClinicalSupport = "";

                                varContactCaseManger = ""; //[U052] - [QDC status update]

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_individualfunding"))
                                    dbIndividualFunding = c.FormattedValues["new_palmclientsupport1.new_individualfunding"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_individualfunding"))
                                    dbIndividualFunding = c.Attributes["new_palmclientsupport1.new_individualfunding"].ToString();
                                else
                                    dbIndividualFunding = "";

                                //Reset case plan information
                                dbCPStartDate = "";
                                dbReviewDate = "";
                                dbCompleted = "";

                                //Loop through the case plan information and get the correct one for the support period
                                foreach (var p in result9.Entities)
                                {
                                    // We need to get the entity id for the client field for comparisons
                                    if (p.Attributes.Contains("new_supportperiod"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)p.Attributes["new_supportperiod"];
                                        dbSupportPeriod = getEntity.Id.ToString();
                                    }
                                    else if (p.FormattedValues.Contains("new_supportperiod"))
                                        dbSupportPeriod = p.FormattedValues["new_supportperiod"];
                                    else
                                        dbSupportPeriod = "";

                                    // Only process data if the support periods match
                                    if (dbSupportPeriod == dbPalmClientSupportId)
                                    {
                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (p.FormattedValues.Contains("new_startdate"))
                                            dbCPStartDate = p.FormattedValues["new_startdate"];
                                        else if (c.Attributes.Contains("new_startdate"))
                                            dbCPStartDate = p.Attributes["new_startdate"].ToString();
                                        else
                                            dbCPStartDate = "";

                                        // Convert date from American format to Australian format
                                        dbCPStartDate = cleanDateAM(dbCPStartDate);

                                        if (p.FormattedValues.Contains("new_reviewdate"))
                                            dbReviewDate = p.FormattedValues["new_reviewdate"];
                                        else if (c.Attributes.Contains("new_reviewdate"))
                                            dbReviewDate = p.Attributes["new_reviewdate"].ToString();
                                        else
                                            dbReviewDate = "";

                                        // Convert date from American format to Australian format
                                        dbReviewDate = cleanDateAM(dbReviewDate);

                                        if (p.FormattedValues.Contains("new_completed"))
                                            dbCompleted = p.FormattedValues["new_completed"];
                                        else if (c.Attributes.Contains("new_completed"))
                                            dbCompleted = p.Attributes["new_completed"].ToString();
                                        else
                                            dbCompleted = "";

                                        if (dbCompleted == "Completed")
                                            dbCompleted = "Yes";
                                        else if (dbCompleted == "None")
                                            dbCompleted = "None";
                                        else if (dbCompleted == "Less Than Half" || dbCompleted == "Half" || dbCompleted == "More Than Half")
                                            dbCompleted = "Some";

                                        break;

                                    } // Same support period
                                } // Case Plan Loop

                                //Set the expected data fields to the case plan start, review and completed dates
                                varPPCreateDateAccOutreach = dbCPStartDate;
                                varPPCreateDateCO = dbCPStartDate;
                                varPPCreateDateCongCare = dbCPStartDate;
                                varPPCreateDateDP = dbCPStartDate;
                                varPPCreateDateFamOpt = dbCPStartDate;
                                varPPCreateDateFFYA = dbCPStartDate;
                                varPPCreateDateHomeFirst = dbCPStartDate;
                                varPPCreateDateISP = dbCPStartDate;
                                varPPCreateDateMA = dbCPStartDate;
                                varPPCreateDatePDSS = dbCPStartDate;
                                varPPCreateDateSharedSup = dbCPStartDate;
                                varPPReviewDateAccOutreach = dbReviewDate;
                                varPPReviewDateCO = dbReviewDate;
                                varPPReviewDateCongCare = dbReviewDate;
                                varPPReviewDateDP = dbReviewDate;
                                varPPReviewDateFamOpt = dbReviewDate;
                                varPPReviewDate_FFYA = dbReviewDate;
                                varPPReviewDateHomeFirst = dbReviewDate;
                                varPPReviewDateISP = dbReviewDate;
                                varPPReviewDateMA = dbReviewDate;
                                varPPReviewDatePDSS = dbReviewDate;
                                varPPReviewDateSharedSup = dbReviewDate;
                                varPPGoals = dbCompleted;

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_clienteft"))
                                    dbClientEft = c.FormattedValues["new_palmclientsupport1.new_clienteft"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_clienteft"))
                                    dbClientEft = c.Attributes["new_palmclientsupport1.new_clienteft"].ToString();
                                else
                                    dbClientEft = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_referraldate"))
                                    dbReferralDate = c.FormattedValues["new_palmclientsupport1.new_referraldate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_referraldate"))
                                    dbReferralDate = c.Attributes["new_palmclientsupport1.new_referraldate"].ToString();
                                else
                                    dbReferralDate = "";

                                // Convert date from American format to Australian format
                                dbReferralDate = cleanDateAM(dbReferralDate);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_sourceref"))
                                    dbSourceRef = c.FormattedValues["new_palmclientsupport1.new_sourceref"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_sourceref"))
                                    dbSourceRef = c.Attributes["new_palmclientsupport1.new_sourceref"].ToString();
                                else
                                    dbSourceRef = "";

                                varReasonRespite = ""; //[U062] - [QDC status update]
                                varNightsReceived_PDSSRespite = ""; //[U063] - [QDC status update]
                                varNightsReceived_PDSSResiRehab = ""; //[U064] - [QDC status update]

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_clientstatus"))
                                    dbClientStatus = c.FormattedValues["new_palmclientsupport1.new_clientstatus"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_clientstatus"))
                                    dbClientStatus = c.Attributes["new_palmclientsupport1.new_clientstatus"].ToString();
                                else
                                    dbClientStatus = "";

                                varLastAssessmentDate = ""; //[U066] - [QDC status update]
                                varNbrMealsAtCentre = ""; //[U067] - [QDC status update]
                                varNbrMealsAtHome = ""; //[U068] - [QDC status update]

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_cessation"))
                                    dbCessation = c.FormattedValues["new_palmclientsupport1.new_cessation"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_cessation"))
                                    dbCessation = c.Attributes["new_palmclientsupport1.new_cessation"].ToString();
                                else
                                    dbCessation = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_startdate"))
                                    dbStartDate = c.FormattedValues["new_palmclientsupport1.new_startdate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_startdate"))
                                    dbStartDate = c.Attributes["new_palmclientsupport1.new_startdate"].ToString();
                                else
                                    dbStartDate = "";

                                // Convert date from American format to Australian format
                                dbStartDate = cleanDateAM(dbStartDate);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_enddate"))
                                    dbEndDate = c.FormattedValues["new_palmclientsupport1.new_enddate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_enddate"))
                                    dbEndDate = c.Attributes["new_palmclientsupport1.new_enddate"].ToString();
                                else
                                    dbEndDate = "";

                                // Convert date from American format to Australian format
                                dbEndDate = cleanDateAM(dbEndDate);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_doqdc"))
                                    dbDoQdc = c.FormattedValues["new_palmclientsupport1.new_doqdc"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_doqdc"))
                                    dbDoQdc = c.Attributes["new_palmclientsupport1.new_doqdc"].ToString();
                                else
                                    dbDoQdc = "";

                                // Loop through QDC agencies to get agency id code
                                foreach (var a in result6.Entities)
                                {
                                    if (a.FormattedValues.Contains("new_agencyid"))
                                        dbAgencyId = a.FormattedValues["new_agencyid"];
                                    else if (a.Attributes.Contains("new_agencyid"))
                                        dbAgencyId = a.Attributes["new_agencyid"].ToString();
                                    else
                                        dbAgencyId = "";

                                    if (a.FormattedValues.Contains("new_agency"))
                                        dbAgency = a.FormattedValues["new_agency"];
                                    else if (a.Attributes.Contains("new_agency"))
                                        dbAgency = a.Attributes["new_agency"].ToString();
                                    else
                                        dbAgency = "";
                                    
                                    if (dbDoQdc.ToLower() == dbAgency.ToLower())
                                    {
                                        dbDoQdc = dbAgencyId;
                                        break;
                                    }
                                }
                                
                                if (c.FormattedValues.Contains("new_palmsuqdcagency2.new_outletcode"))
                                    dbOutletCode2 = c.FormattedValues["new_palmsuqdcagency2.new_outletcode"];
                                else if (c.Attributes.Contains("new_palmsuqdcagency2.new_outletcode"))
                                    dbOutletCode2 = c.GetAttributeValue<AliasedValue>("new_palmsuqdcagency2.new_outletcode").Value.ToString();
                                else
                                    dbOutletCode2 = "";

                                //Get the agency value based on the QDC code
                                varAgencyValue = 0;

                                if (String.IsNullOrEmpty(dbOutletCode2) == true)
                                    dbOutletCode2 = "00";

                                if (dbOutletCode2.Length > 1)
                                {
                                    if (dbOutletCode2.Substring(0, 2) == "17") //DSD
                                        varAgencyValue = 1;
                                    else if (dbOutletCode2.Substring(0, 2) == "15") //PDRSS
                                        varAgencyValue = 2;
                                    else if (dbOutletCode2.Substring(0, 2) == "13") //HACC
                                        varAgencyValue = 3;
                                    else
                                    {
                                        varAgencyValue = 0;
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Invalid QDC code<br>");
                                    }
                                }
                                else
                                {
                                    varAgencyValue = 0;
                                    sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Invalid QDC code<br>");
                                }

                                // Get the outlet code number
                                varService_Type_Outlet_Codeno = dbDoQdc;

                                // Validate the postcode or apply default data
                                if (String.IsNullOrEmpty(dbLocality) == false && String.IsNullOrEmpty(dbPostcode) == false && String.IsNullOrEmpty(dbState) == false)
                                {
                                    //Postcode needs to be 4 digits
                                    if (dbPostcode.Length == 1)
                                        dbPostcode = "000" + dbPostcode;
                                    else if (dbPostcode.Length == 2)
                                        dbPostcode = "00" + dbPostcode;
                                    else if (dbPostcode.Length == 3)
                                        dbPostcode = "0" + dbPostcode;
                                    else if (dbPostcode.Length > 4)
                                        dbPostcode = "9999";
                                }
                                else
                                {
                                    dbLocality = "Unknown";
                                    dbPostcode = "";
                                    dbState = "2";
                                }

                                // Loop through drop down list data
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

                                    if (d.FormattedValues.Contains("new_qdc"))
                                        varQDC = d.FormattedValues["new_qdc"];
                                    else if (d.Attributes.Contains("new_qdc"))
                                        varQDC = d.Attributes["new_qdc"].ToString();
                                    else
                                        varQDC = "";

                                    // Ensure numeric
                                    varQDC = cleanString(varQDC, "number");
                                    if (String.IsNullOrEmpty(varQDC) == true)
                                        varQDC = "0";

                                    // Check the value against the description for the correct data type, and assign the QDC number
                                    if (varType.ToLower() == "crship")
                                    {
                                        if (dbCarerRship.ToLower() == varDesc.ToLower())
                                            dbCarerRship = varQDC;
                                    }

                                    if (varType.ToLower() == "carerage")
                                    {
                                        if (dbCarerDob.ToLower() == varDesc.ToLower())
                                            dbCarerDob = varQDC;
                                    }

                                    if (varType.ToLower() == "communication")
                                    {
                                        if (dbCommunication.ToLower() == varDesc.ToLower())
                                            dbCommunication = varQDC;
                                    }

                                    if (varType.ToLower() == "indigenous")
                                    {
                                        if (dbIndigenous.ToLower() == varDesc.ToLower())
                                            dbIndigenous = varQDC;
                                    }

                                    if (varType.ToLower() == "ccpdis")
                                    {
                                        if (dbCCPDisType.ToLower() == varDesc.ToLower())
                                            dbCCPDisType = varQDC;
                                    }

                                    if (varType.ToLower() == "diagnosis")
                                    {
                                        if (dbPrimaryDiagnosis.ToLower() == varDesc.ToLower())
                                            dbPrimaryDiagnosis = varQDC;
                                    }

                                    if (varType.ToLower() == "livingarrange")
                                    {
                                        if (dbLivingArrangePres.ToLower() == varDesc.ToLower())
                                            dbLivingArrangePres = varQDC;

                                        if (dbLivesWithAtExit.ToLower() == varDesc.ToLower())
                                            dbLivesWithAtExit = varQDC;
                                    }

                                    if (varType.ToLower() == "ccpdepchild")
                                    {
                                        if (dbDepChild.ToLower() == varDesc.ToLower())
                                            dbDepChild = varQDC;
                                    }

                                    if (varType.ToLower() == "residential")
                                    {
                                        if (dbResidentialPres.ToLower() == varDesc.ToLower())
                                            dbResidentialPres = varQDC;

                                        if (dbAccommodationExit.ToLower() == varDesc.ToLower())
                                            dbAccommodationExit = varQDC;
                                    }

                                    if (varType.ToLower() == "labour")
                                    {
                                        if (dbLabourForcePres.ToLower() == varDesc.ToLower())
                                            dbLabourForcePres = varQDC;
                                    }

                                    if (varType.ToLower() == "income")
                                    {
                                        if (dbIncomePres.ToLower() == varDesc.ToLower())
                                            dbIncomePres = varQDC;
                                    }

                                    if (varType.ToLower() == "clinicalsupport")
                                    {
                                        if (dbClinicalSupport.ToLower() == varDesc.ToLower())
                                            dbClinicalSupport = varQDC;
                                    }

                                    if (varType.ToLower() == "sourceref")
                                    {
                                        if (dbSourceRef.ToLower() == varDesc.ToLower())
                                            dbSourceRef = varQDC;
                                    }

                                    if (varType.ToLower() == "cessation")
                                    {
                                        if (dbCessation.ToLower() == varDesc.ToLower())
                                            dbCessation = varQDC;
                                    }

                                    if (varType.ToLower() == "state")
                                    {
                                        if (dbState.ToLower() == varDesc.ToLower())
                                            dbState = varQDC;
                                    }

                                } // Drop List

                                // Country value for don't know / not applicable
                                if (dbCountry == "Don't Know" || dbCountry == "Not Applicable")
                                    dbCountry = "9999";

                                //Get the surname part of the SLK
                                if (String.IsNullOrEmpty(dbSurname) == false)
                                {
                                    dbSurname = removeString(dbSurname, "html");
                                    varSurname = cleanString(dbSurname.ToUpper(), "letter");
                                }
                                else
                                    varSurname = "999999";

                                varSurname = varSurname + "22222";
                                varSurname = varSurname.Substring(1, 1) + varSurname.Substring(2, 1) + varSurname.Substring(4, 1);

                                if (varSurname == "222")
                                    varSurname = "999";

                                //Get the firstname part of the SLK
                                if (String.IsNullOrEmpty(dbFirstName) == false)
                                {
                                    dbFirstName = removeString(dbFirstName, "html");
                                    varFirstName = cleanString(dbFirstName.ToUpper(), "letter");
                                }
                                else
                                    varFirstName = "999999";

                                varFirstName = varFirstName + "22222";
                                varFirstName = varFirstName.Substring(1, 1) + varFirstName.Substring(2, 1);

                                if (varFirstName == "22")
                                    varFirstName = "99";

                                //Get the gender code
                                if (dbGender == "Female")
                                    dbGender = "2";
                                else
                                    dbGender = "1";

                                //Get the estimated flag
                                if (dbDobEst == "Yes")
                                    dbDobEst = "Y";
                                else if (dbDobEst == "No")
                                    dbDobEst = "N";
                                else
                                    sbErrorList.AppendLine("Data Missing: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - DOB Estimate Flag<br>");

                                //Check for valid date of birth
                                if (String.IsNullOrEmpty(dbDob) == false)
                                {
                                    DateTime.TryParse(dbDob, out varCheckDate);

                                    //Get the age based on the date of birth
                                    varAge = varCurrentDate.Year - varCheckDate.Year;
                                    if (varCurrentDate.Month < varCheckDate.Month || (varCurrentDate.Month == varCheckDate.Month && varCurrentDate.Day < varCheckDate.Day))
                                        varAge--;

                                    if (varCheckDate.Year <= (varCurrentDate.Year - 115))
                                    {
                                        dbDob = "1-Jan-1970";
                                        dbDobEst = "Y";
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - DOB<br>");
                                        varAge = 0;
                                    }

                                }
                                else
                                {
                                    dbDob = "1-Jan-1970";
                                    dbDobEst = "Y";
                                    varAge = 0;
                                }
                                                                

                                //Put dob into expected format
                                varDob = cleanDateS(varCheckDate);

                                //Always consent
                                dbConsent = "1";

                                //Get number based on value or alert of error
                                if (dbCarerAvail == "Has A Carer")
                                    dbCarerAvail = "1";
                                else if (dbCarerAvail == "Has No Carer")
                                    dbCarerAvail = "2";
                                else
                                {
                                    sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Existance of Carer not defined<br>");
                                    dbCarerAvail = "0";
                                }

                                //Create the statistical linkage key
                                varSLK = varSurname + varFirstName + varDob + dbGender;

                                //For DSD and PDRSS only
                                if (varAgencyValue == 1 || varAgencyValue == 2)
                                {
                                    //If client has carer, get number based on value or alert of error
                                    if (dbCarerAvail == "1")
                                    {
                                        if (dbCarerPrimary == "Yes")
                                            dbCarerPrimary = "1";
                                        else if (dbCarerPrimary == "No")
                                            dbCarerPrimary = "2";
                                        else
                                            sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Carer primary status not defined<br>");
                                    }
                                    //No carer
                                    else
                                        dbCarerPrimary = "";

                                }


                                //For DSD, PDRSS, and HACC
                                if (varAgencyValue == 1 || varAgencyValue == 2 || varAgencyValue == 3)
                                {

                                    //If client has carer, get number based on value or alert of error
                                    if (dbCarerAvail == "1")
                                    {
                                        if (dbCarerResidence == "Co-resident carer")
                                            dbCarerResidence = "1";
                                        else if (dbCarerResidence == "Non Resident Carer" || dbCarerResidence == "Not stated/inadequately described" || dbCarerResidence == "Not Applicable")
                                            dbCarerResidence = "2";
                                        else
                                            sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Carer's residence status not defined<br>");

                                        //If not numeric then value was not found in drop down list
                                        if (!Int32.TryParse(dbCarerRship, out varCheckInt))
                                            sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Carer relationship not defined<br>");

                                    }
                                    //No carer
                                    else
                                    {
                                        dbCarerResidence = "";
                                        dbCarerRship = "";
                                    }

                                }


                                //For DSD and PDRSS only
                                if (varAgencyValue == 1 || varAgencyValue == 2)
                                {

                                    //If client has carer, get number based on value or alert of error
                                    if (dbCarerAvail == "1")
                                    {
                                        varCAge = 0;
                                        if (String.IsNullOrEmpty(dbCarerDob) == false)
                                        {
                                            DateTime.TryParse(dbCarerDob, out varCheckDate);
                                            varCAge = varCurrentDate.Year - varCheckDate.Year;

                                            if (varCurrentDate.Month < varCheckDate.Month || (varCurrentDate.Month == varCheckDate.Month && varCurrentDate.Day < varCheckDate.Day))
                                                varCAge--;

                                            if (varCAge < 15)
                                                dbCarerDob = "1";
                                            else if (varCAge >= 15 && varCAge <= 24)
                                                dbCarerDob = "2";
                                            else if (varCAge >= 25 && varCAge <= 44)
                                                dbCarerDob = "3";
                                            else if (varCAge >= 45 && varCAge <= 64)
                                                dbCarerDob = "4";
                                            else
                                                dbCarerDob = "5";

                                        }
                                        else
                                            sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Carer age group not defined<br>");

                                    }
                                    //No carer
                                    else
                                        dbCarerDob = "";

                                }


                                //For PDRSS only
                                if (varAgencyValue == 2)
                                {

                                    //If client has carer, get number based on value or alert of error
                                    if (dbCarerAvail == "1")
                                    {

                                        if (dbCarerPdrssClient == "Yes - as a primary client")
                                            dbCarerPdrssClient = "1";
                                        else if (dbCarerPdrssClient == "Yes - as a carer of a client")
                                            dbCarerPdrssClient = "2";
                                        else if (dbCarerPdrssClient == "No")
                                            dbCarerPdrssClient = "4";
                                        else
                                            sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Carer as a PDRSS client not defined<br>");

                                    }
                                    //No carer
                                    else
                                        dbCarerPdrssClient = "";

                                }


                                //For DSD, PDRSS, and HACC
                                if (varAgencyValue == 1 || varAgencyValue == 2 || varAgencyValue == 3)
                                {
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbIndigenous, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Indigenous status not defined<br>");
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbCountry, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Country of birth not defined<br>");
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbLanguage, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Client language not defined<br>");

                                }


                                //For DSD and PDRSS only
                                if (varAgencyValue == 1 || varAgencyValue == 2)
                                {

                                    //Get number based on value or alert of error
                                    if (dbInterpret == "Yes, for spoken language other than English")
                                        dbInterpret = "1";
                                    else if (dbInterpret == "Yes, for non-spoken communication")
                                        dbInterpret = "2";
                                    else if (dbInterpret == "No" || dbInterpret == "Not stated")
                                        dbInterpret = "3";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Need for interpreter not defined<br>");

                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbCommunication, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Communication method not defined<br>");

                                }


                                //For DSD only
                                if (varAgencyValue == 1)
                                {

                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbCCPDisType, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Primary disability group not defined<br>");
                                }


                                //For PDRSS only
                                if (varAgencyValue == 2)
                                {

                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbPrimaryDiagnosis, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Primary diagnosis not defined<br>");
                                }


                                //For specific agency code
                                if ((varAgencyValue == 1 || varAgencyValue == 2 || varAgencyValue == 3) && dbOutletCode2 != "15038")
                                {
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbLivingArrangePres, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Living arrangements not defined<br>");
                                }


                                //For PDRSS only
                                if (varAgencyValue == 2)
                                {
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbDepChild, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Dependant children not defined<br>");
                                }


                                //For specific agency codes only
                                if (dbOutletCode2 == "15034" || dbOutletCode2 == "15038" || dbOutletCode2 == "15055")
                                {
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbLivingArrangePres, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Living arrangements at entry not defined<br>");
                                    //Only do for exited client
                                    if (String.IsNullOrEmpty(dbEndDate) == false)
                                    {
                                        //If not numeric then value was not found in drop down list
                                        if (!Int32.TryParse(dbLivesWithAtExit, out varCheckInt))
                                            sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Living arrangements at exit not defined<br>");
                                    }
                                }


                                //For DSD, PDRSS and HACC
                                if (varAgencyValue == 1 || varAgencyValue == 2 || varAgencyValue == 3)
                                {
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbResidentialPres, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Residential setting not defined<br>");
                                }


                                //For specific agency codes only
                                if (dbOutletCode2 == "15034" || dbOutletCode2 == "15038" || dbOutletCode2 == "15055")
                                {

                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbResidentialPres, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Residential setting prior to entry not defined<br>");

                                    //Only do for exited client
                                    if (String.IsNullOrEmpty(dbEndDate) == false)
                                    {
                                        //If not numeric then value was not found in drop down list
                                        if (!Int32.TryParse(dbAccommodationExit, out varCheckInt))
                                            sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Residential setting at exit not defined<br>");
                                    }

                                }


                                //For HACC only
                                if (varAgencyValue == 3)
                                {
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbAccommodationExit, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Accommodation after cessation not defined<br>");
                                }


                                //For specific agency only
                                if (dbOutletCode2 == "15034")
                                {
                                    //Get number based on value
                                    if (dbNominationRights == "Yes")
                                        dbNominationRights = "1";
                                    else if (dbNominationRights == "No")
                                        dbNominationRights = "2";
                                    else
                                        dbNominationRights = "";

                                }


                                //For DSD and PDRSS only
                                if (varAgencyValue == 1 || varAgencyValue == 2)
                                {

                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbLabourForcePres, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Labour force status not defined<br>");

                                    //For clients older than 16
                                    if (varAge >= 16)
                                    {
                                        //If not numeric then value was not found in drop down list
                                        if (!Int32.TryParse(dbIncomePres, out varCheckInt))
                                            sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Income source not defined<br>");
                                    }
                                    //Client not older than 16
                                    else
                                        dbIncomePres = "";

                                    //For clients younder than 16
                                    if (varAge < 16)
                                    {
                                        //Get number based on value or alert of error
                                        if (dbCarerAllowance == "Yes")
                                            dbCarerAllowance = "1";
                                        else if (dbCarerAllowance == "No")
                                            dbCarerAllowance = "2";
                                        else if (dbCarerAllowance == "Not Known")
                                            dbCarerAllowance = "97";
                                        else
                                            sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Receipt of carer allowance not defined<br>");

                                    }
                                    //Client not younger than 16
                                    else
                                        dbCarerAllowance = "";

                                }


                                //For DSD and PDRSS only
                                if (varAgencyValue == 2 || varAgencyValue == 3)
                                {

                                    //Get the benefit status based on the income source
                                    if (dbIncomePres == "1")
                                        varPensionBenefitStatus = "1";
                                    else if (dbIncomePres == "2")
                                        varPensionBenefitStatus = "6";
                                    else if (dbIncomePres == "3")
                                        varPensionBenefitStatus = "4";
                                    else if (dbIncomePres == "4")
                                        varPensionBenefitStatus = "7";
                                    else if (dbIncomePres == "5")
                                        varPensionBenefitStatus = "6";
                                    else if (dbIncomePres == "6")
                                        varPensionBenefitStatus = "3";
                                    else if (dbIncomePres == "7")
                                        varPensionBenefitStatus = "7";
                                    else if (dbIncomePres == "8")
                                        varPensionBenefitStatus = "5";
                                    else if (dbIncomePres == "9")
                                        varPensionBenefitStatus = "6";
                                    else if (dbIncomePres == "10")
                                        varPensionBenefitStatus = "6";
                                    else if (dbIncomePres == "11")
                                        varPensionBenefitStatus = "6";
                                    else if (dbIncomePres == "12")
                                        varPensionBenefitStatus = "7";
                                    else if (dbIncomePres == "13")
                                        varPensionBenefitStatus = "7";
                                    else if (dbIncomePres == "14")
                                        varPensionBenefitStatus = "6";
                                    else if (dbIncomePres == "15")
                                        varPensionBenefitStatus = "6";
                                    else if (dbIncomePres == "16")
                                        varPensionBenefitStatus = "7";
                                    else if (dbIncomePres == "17")
                                        varPensionBenefitStatus = "7";
                                    else if (dbIncomePres == "18")
                                        varPensionBenefitStatus = "7";
                                    else
                                        varPensionBenefitStatus = "10";

                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(varPensionBenefitStatus, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Pension / benefit status not defined<br>");

                                }


                                //For DSD only
                                if (varAgencyValue == 1)
                                {

                                    //Get number based on value or alert of error
                                    if (dbOutsideNoTransport == "Fully")
                                        dbOutsideNoTransport = "1";
                                    else if (dbOutsideNoTransport == "Partially")
                                        dbOutsideNoTransport = "2";
                                    else if (dbOutsideNoTransport == "Not at all")
                                        dbOutsideNoTransport = "3";
                                    else if (dbOutsideNoTransport == "Not known")
                                        dbOutsideNoTransport = "97";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Participate Outside No Transport not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbParticipateTransport == "Fully")
                                        dbParticipateTransport = "1";
                                    else if (dbParticipateTransport == "Partially")
                                        dbParticipateTransport = "2";
                                    else if (dbParticipateTransport == "Not at all")
                                        dbParticipateTransport = "3";
                                    else if (dbParticipateTransport == "Not known")
                                        dbParticipateTransport = "97";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Participate Transport not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbParticipateFamily == "Fully")
                                        dbParticipateFamily = "1";
                                    else if (dbParticipateFamily == "Partially")
                                        dbParticipateFamily = "2";
                                    else if (dbParticipateFamily == "Not at all")
                                        dbParticipateFamily = "3";
                                    else if (dbParticipateFamily == "Not known")
                                        dbParticipateFamily = "97";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Participate Family not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbParticipateSocial == "Fully")
                                        dbParticipateSocial = "1";
                                    else if (dbParticipateSocial == "Partially")
                                        dbParticipateSocial = "2";
                                    else if (dbParticipateSocial == "Not at all")
                                        dbParticipateSocial = "3";
                                    else if (dbParticipateSocial == "Not known")
                                        dbParticipateSocial = "97";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Participate Social not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbParticipateLeisure == "Fully")
                                        dbParticipateLeisure = "1";
                                    else if (dbParticipateLeisure == "Partially")
                                        dbParticipateLeisure = "2";
                                    else if (dbParticipateLeisure == "Not at all")
                                        dbParticipateLeisure = "3";
                                    else if (dbParticipateLeisure == "Not known")
                                        dbParticipateLeisure = "97";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Participate Leisure not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbParticipateWorking == "Fully")
                                        dbParticipateWorking = "1";
                                    else if (dbParticipateWorking == "Partially")
                                        dbParticipateWorking = "2";
                                    else if (dbParticipateWorking == "Not at all")
                                        dbParticipateWorking = "3";
                                    else if (dbParticipateWorking == "Not known")
                                        dbParticipateWorking = "97";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Participate Working not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbParticipateMoney == "Fully")
                                        dbParticipateMoney = "1";
                                    else if (dbParticipateMoney == "Partially")
                                        dbParticipateMoney = "2";
                                    else if (dbParticipateMoney == "Not at all")
                                        dbParticipateMoney = "3";
                                    else if (dbParticipateMoney == "Not known")
                                        dbParticipateMoney = "97";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Participate Money not defined<br>");

                                } //DSD only


                                //For DSD and PDRSS only
                                if (varAgencyValue == 1 || varAgencyValue == 2)
                                {

                                    //Get number based on value or alert of error
                                    if (dbSupportSelfCare == "Always needs help")
                                        dbSupportSelfCare = "1";
                                    else if (dbSupportSelfCare == "Sometimes needs help")
                                        dbSupportSelfCare = "2";
                                    else if (dbSupportSelfCare == "Uses aids")
                                        dbSupportSelfCare = "3";
                                    else if (dbSupportSelfCare == "Does not need help")
                                        dbSupportSelfCare = "4";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Support Self Care not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbSupportMobility == "Always needs help")
                                        dbSupportMobility = "1";
                                    else if (dbSupportMobility == "Sometimes needs help")
                                        dbSupportMobility = "2";
                                    else if (dbSupportMobility == "Uses aids")
                                        dbSupportMobility = "3";
                                    else if (dbSupportMobility == "Does not need help")
                                        dbSupportMobility = "4";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Support Mobility not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbSupportCommunication == "Always needs help")
                                        dbSupportCommunication = "1";
                                    else if (dbSupportCommunication == "Sometimes needs help")
                                        dbSupportCommunication = "2";
                                    else if (dbSupportCommunication == "Uses aids")
                                        dbSupportCommunication = "3";
                                    else if (dbSupportCommunication == "Does not need help")
                                        dbSupportCommunication = "4";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Support Communication not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbSupportInterpersonal == "Always needs help")
                                        dbSupportInterpersonal = "1";
                                    else if (dbSupportInterpersonal == "Sometimes needs help")
                                        dbSupportInterpersonal = "2";
                                    else if (dbSupportInterpersonal == "Uses aids")
                                        dbSupportInterpersonal = "3";
                                    else if (dbSupportInterpersonal == "Does not need help")
                                        dbSupportInterpersonal = "4";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Support Interpersonal not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbSupportLearning == "Always needs help")
                                        dbSupportLearning = "1";
                                    else if (dbSupportLearning == "Sometimes needs help")
                                        dbSupportLearning = "2";
                                    else if (dbSupportLearning == "Uses aids")
                                        dbSupportLearning = "3";
                                    else if (dbSupportLearning == "Does not need help")
                                        dbSupportLearning = "4";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Support Learning not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbSupportEducation == "Always needs help")
                                        dbSupportEducation = "1";
                                    else if (dbSupportEducation == "Sometimes needs help")
                                        dbSupportEducation = "2";
                                    else if (dbSupportEducation == "Uses aids")
                                        dbSupportEducation = "3";
                                    else if (dbSupportEducation == "Does not need help")
                                        dbSupportEducation = "4";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Support Education not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbSupportCommunity == "Always needs help")
                                        dbSupportCommunity = "1";
                                    else if (dbSupportCommunity == "Sometimes needs help")
                                        dbSupportCommunity = "2";
                                    else if (dbSupportCommunity == "Uses aids")
                                        dbSupportCommunity = "3";
                                    else if (dbSupportCommunity == "Does not need help")
                                        dbSupportCommunity = "4";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Support Community not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbSupportDomestic == "Always needs help")
                                        dbSupportDomestic = "1";
                                    else if (dbSupportDomestic == "Sometimes needs help")
                                        dbSupportDomestic = "2";
                                    else if (dbSupportDomestic == "Uses aids")
                                        dbSupportDomestic = "3";
                                    else if (dbSupportDomestic == "Does not need help")
                                        dbSupportDomestic = "4";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Support Domestic not defined<br>");

                                    //Get number based on value or alert of error
                                    if (dbSupportWorking == "Always needs help")
                                        dbSupportWorking = "1";
                                    else if (dbSupportWorking == "Sometimes needs help")
                                        dbSupportWorking = "2";
                                    else if (dbSupportWorking == "Uses aids")
                                        dbSupportWorking = "3";
                                    else if (dbSupportWorking == "Does not need help")
                                        dbSupportWorking = "4";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Support Working not defined<br>");

                                } //DSD and PDRSS


                                //For PDRSS only
                                if (varAgencyValue == 2)
                                {

                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbClinicalSupport, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Clinical Support not defined<br>");

                                    //Get number based on value
                                    if (varContactCaseManger == "Yes")
                                        varContactCaseManger = "1";
                                    else if (varContactCaseManger == "No")
                                        varContactCaseManger = "2";
                                    else if (varContactCaseManger == "Not known")
                                        varContactCaseManger = "97";
                                    else
                                        varContactCaseManger = "";
                                    //[Warn Fix]
                                    //sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Contact Case Manager not defined<br>");

                                }


                                //For DSD and PDRSS only
                                if (varAgencyValue == 1 || varAgencyValue == 2)
                                {

                                    //Get number based on value
                                    if (dbIndividualFunding == "Yes")
                                        dbIndividualFunding = "1";
                                    else if (dbIndividualFunding == "No")
                                        dbIndividualFunding = "2";
                                    else if (dbIndividualFunding == "Not known")
                                        dbIndividualFunding = "97";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Individual Funding Status not defined<br>");

                                }

                                //For PDRSS Only
                                if (varAgencyValue == 2)
                                {

                                    //Get number based on value
                                    if (varPPGoals == "Yes")
                                        varPPGoals = "1";
                                    else if (varPPGoals == "None")
                                        varPPGoals = "2";
                                    else if (varPPGoals == "Some")
                                        varPPGoals = "3";
                                    else
                                        varPPGoals = "";

                                    //[Warn Fix]
                                    //sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - PP Goals not defined<br>");

                                }


                                //For specific agency only
                                if (dbOutletCode2 == "17022")
                                {
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbClientEft, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Client EFT not defined<br>");
                                }


                                //Check for referral date based on agency id
                                if (dbOutletCode2 == "17006")
                                {
                                    if (String.IsNullOrEmpty(dbReferralDate) == true)
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Referral Date not defined<br>");
                                }
                                if (dbOutletCode2 == "17026")
                                {
                                    if (String.IsNullOrEmpty(dbReferralDate) == true)
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Referral Date not defined<br>");
                                }
                                if (dbOutletCode2 == "17028")
                                {
                                    if (String.IsNullOrEmpty(dbReferralDate) == true)
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Referral Date not defined<br>");
                                }
                                if (dbOutletCode2 == "15034")
                                {
                                    if (String.IsNullOrEmpty(dbReferralDate) == true)
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Referral Date not defined<br>");
                                }
                                if (dbOutletCode2 == "15037")
                                {
                                    if (String.IsNullOrEmpty(dbReferralDate) == true)
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Referral Date not defined<br>");
                                }
                                if (dbOutletCode2 == "15035")
                                {
                                    if (String.IsNullOrEmpty(dbReferralDate) == true)
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Referral Date not defined<br>");
                                }
                                if (dbOutletCode2 == "15038")
                                {
                                    if (String.IsNullOrEmpty(dbReferralDate) == true)
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Referral Date not defined<br>");
                                }
                                if (dbOutletCode2 == "15055")
                                {
                                    if (String.IsNullOrEmpty(dbReferralDate) == true)
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Referral Date not defined<br>");
                                }
                                if (dbOutletCode2 == "17042")
                                {
                                    if (String.IsNullOrEmpty(dbReferralDate) == true)
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Referral Date not defined<br>");
                                }


                                //For PDRSS only
                                if (varAgencyValue == 2)
                                {
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbSourceRef, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Referral source not defined<br>");
                                }


                                //For HACC only
                                if (varAgencyValue == 3)
                                {
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbSourceRef, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Referral source not defined<br>");
                                }


                                //For specific agency code only
                                if (dbOutletCode2 == "15037")
                                {

                                    //Get number based on value or alert of error
                                    if (varReasonRespite == "Emergency (non clinical)")
                                        varReasonRespite = "1";
                                    else if (varReasonRespite == "Planned")
                                        varReasonRespite = "2";
                                    else if (varReasonRespite == "Other")
                                        varReasonRespite = "3";
                                    else
                                        varReasonRespite = "";
                                    //[Warn Fix]
                                    //sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Reason for Respite not defined<br>");

                                    //Must be numeric
                                    if (!Int32.TryParse(varNightsReceived_PDSSRespite, out varCheckInt))
                                        varNightsReceived_PDSSRespite = "";
                                    //[Warn Fix]
                                    //sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Nights of Respite not defined<br>");

                                }


                                //For specific agency codes only
                                if (dbOutletCode2 == "15038" || dbOutletCode2 == "15055")
                                {
                                    //Must be numeric
                                    if (!Int32.TryParse(varNightsReceived_PDSSResiRehab, out varCheckInt))
                                        varNightsReceived_PDSSResiRehab = "";
                                    //[Warn Fix]
                                    //sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Nights in Residential Rehabilitation not defined<br>");
                                }


                                //For HACC only
                                if (varAgencyValue == 3)
                                {

                                    //Get number based on value or alert of error
                                    if (dbClientStatus == "This record describes the client")
                                        dbClientStatus = "1";
                                    else if (dbClientStatus == "This record describes the carer")
                                        dbClientStatus = "3";
                                    else
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Client Reason not defined<br>");

                                    //Must have valid date
                                    if (String.IsNullOrEmpty(varLastAssessmentDate) == true)
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Last Assessment Date not defined<br>");

                                    //Must be numeric
                                    if (!Int32.TryParse(varNbrMealsAtCentre, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Number of meals at centre not defined<br>");

                                    //Must be numeric
                                    if (!Int32.TryParse(varNbrMealsAtHome, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Number of meals at home not defined<br>");

                                }


                                //Start date must be valid
                                if (String.IsNullOrEmpty(dbStartDate) == true)
                                    sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Start Date<br>");

                                //End date must be valid and must not be greater than end of quarter
                                if (String.IsNullOrEmpty(dbEndDate) == false)
                                {
                                    if (Convert.ToDateTime(dbEndDate) > Convert.ToDateTime(varEndDate))
                                        dbEndDate = "";
                                }


                                //For DSD and PDRSS if valid end date
                                if ((varAgencyValue == 1 || varAgencyValue == 2) && String.IsNullOrEmpty(dbEndDate) == false)
                                {
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbCessation, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Exit Reason not defined<br>");
                                }

                                //For HACC if valid end date
                                if (varAgencyValue == 3 && String.IsNullOrEmpty(dbEndDate) == false)
                                {
                                    //If not numeric then value was not found in drop down list
                                    if (!Int32.TryParse(dbCessation, out varCheckInt))
                                        sbErrorList.AppendLine("Invalid Data: (" + dbPalmClientSupportId + ") " + dbFirstName + " " + dbSurname + " - Exit Reason not defined<br>");
                                }



                                varLast_Service_Date = cleanDate(varEndDate); //'The last day of service for the month | If no value set to last day of month
                                varSnapshot_Type = "N"; //'[U071] Was service received on last Wednesday of Financial year

                                //Get the end of the financial year
                                if (varQuarter == 1 || varQuarter == 2)
                                    varSnapShotDay = Convert.ToDateTime("30-Jun-" + (varYear + 1));
                                else
                                    varSnapShotDay = Convert.ToDateTime("30-Jun-" + varYear);


                                //Subtract a day until last Wednesday of the year
                                while (varSnapShotDay.ToString("dddd") != "Wednesday")
                                    varSnapShotDay = varSnapShotDay.AddDays(-1);

                                //Loop through activities to get start date
                                foreach (var p in result8.Entities)
                                {
                                    // We need to get the entity id for the client field for comparisons
                                    if (p.Attributes.Contains("new_supportperiod"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)p.Attributes["new_supportperiod"];
                                        dbSupportPeriod = getEntity.Id.ToString();
                                    }
                                    else if (p.FormattedValues.Contains("new_supportperiod"))
                                        dbSupportPeriod = p.FormattedValues["new_supportperiod"];
                                    else
                                        dbSupportPeriod = "";

                                    if (p.FormattedValues.Contains("new_entrydate"))
                                        dbEntryDate = p.FormattedValues["new_entrydate"];
                                    else if (c.Attributes.Contains("new_entrydate"))
                                        dbEntryDate = p.Attributes["new_entrydate"].ToString();
                                    else
                                        dbEntryDate = "";

                                    // Convert date from American format to Australian format
                                    dbEntryDate = cleanDateAM(dbEntryDate);

                                    // If the support periods are the same, get the service date
                                    if (dbSupportPeriod == dbPalmClientSupportId && string.IsNullOrEmpty(dbEntryDate) == false)
                                    {
                                        //Get the last service date for this client
                                        if (String.IsNullOrEmpty(varLast_Service_Date) == true)
                                            varLast_Service_Date = dbEntryDate;
                                        else if (Convert.ToDateTime(dbEntryDate) > Convert.ToDateTime(varLast_Service_Date))
                                            varLast_Service_Date = dbEntryDate;

                                        //Set flag if at least one activity falls on the snapshot date
                                        if (Convert.ToDateTime(dbEntryDate) == Convert.ToDateTime(varSnapShotDay))
                                            varSnapshot_Type = "Y";
                                    } // Same support period

                                } // Activities Loop

                                //Create second part of extract
                                //Only insert a line if it is specific for this client (as defined above)
                                sbHeaderList.AppendLine("   <service_user postcode=\"" + dbPostcode + "\" has_carer_type=\"" + dbCarerAvail + "\" birth_date=\"" + dbDob + "\" birth_date_est_ind=\"" + dbDobEst + "\" sex_code=\"" + dbGender + "\" stats_link_key=\"" + varSLK + "\" suburb=\"" + dbLocality + "\" state_code=\"" + dbState + "\" consent_type=\"" + dbConsent + "\">");


                                //Equipment. How to handle? Multi select?
                                //If varAgencyValue = 3 AND 1=2 Then
                                //sbHeaderList.AppendLine("      <service_equipment issue_date=\"25 Jul 2002\" equipment_code=\"20\"/>");

                                // Do this if there is a carer
                                if (dbCarerAvail == "1")
                                {

                                    //DSD and PDRSS
                                    if (varAgencyValue == 1 || varAgencyValue == 2)
                                        sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"CarerPrimary\" value=\"" + dbCarerPrimary + "\" />");

                                    if (varAgencyValue == 1 || varAgencyValue == 2 || varAgencyValue == 3)
                                    {
                                        sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"CarerResidence\" value=\"" + dbCarerResidence + "\" />");
                                        sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"CarerRelationship\" value=\"" + dbCarerRship + "\" />");
                                    }

                                    if (varAgencyValue == 1 || varAgencyValue == 2)
                                        sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"CarerAgeGroup\" value=\"" + dbCarerDob + "\" />");

                                    if (varAgencyValue == 2)
                                        sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"CarerPDSSClient\" value=\"" + dbCarerPdrssClient + "\" />");

                                } //Has Carer


                                if (varAgencyValue == 1 || varAgencyValue == 2 || varAgencyValue == 3)
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"IndigenousStatus\" value=\"" + dbIndigenous + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"BirthCountry\" value=\"" + dbCountry + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"Language\" value=\"" + dbLanguage + "\" />");
                                }

                                if (varAgencyValue == 1 || varAgencyValue == 2)
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"InterpreterServices\" value=\"" + dbInterpret + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"CommunicationMethod\" value=\"" + dbCommunication + "\" />");
                                }


                                //Create a multi response list to disability type for DSD clients
                                if (varAgencyValue == 1)
                                {

                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PrimaryDisabilityCSTDA\" value=\"" + dbCCPDisType + "\" />");
                                    // Loop through drop down list values to get QDC value
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

                                        if (d.FormattedValues.Contains("new_qdc"))
                                            varQDC = d.FormattedValues["new_qdc"];
                                        else if (d.Attributes.Contains("new_qdc"))
                                            varQDC = d.Attributes["new_qdc"].ToString();
                                        else
                                            varQDC = "";

                                        varQDC = cleanString(varQDC, "number");
                                        if (String.IsNullOrEmpty(varQDC) == true)
                                            varQDC = "0";

                                        if (varType.ToLower() == "ccpdis" && String.IsNullOrEmpty(dbOtherDis) == false)
                                        {
                                            if (dbOtherDis.IndexOf("*" + varQDC + "*") > -1)
                                                sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"OtherDisabilityCSTDA\" value=\"" + varQDC + "\" />");
                                        }

                                    } //Drop down Loop

                                } //Agency Value = 1


                                //Create a multi response list to disability type for PDRSS clients
                                if (varAgencyValue == 2)
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PrimaryDiagnosisPDSS\" value=\"" + dbPrimaryDiagnosis + "\" />");
                                    // Loop through drop down list values to get QDC value
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

                                        if (d.FormattedValues.Contains("new_qdc"))
                                            varQDC = d.FormattedValues["new_qdc"];
                                        else if (d.Attributes.Contains("new_qdc"))
                                            varQDC = d.Attributes["new_qdc"].ToString();
                                        else
                                            varQDC = "";

                                        varQDC = cleanString(varQDC, "number");
                                        if (String.IsNullOrEmpty(varQDC) == true)
                                            varQDC = "0";

                                        if (varType.ToLower() == "ccpdis" && String.IsNullOrEmpty(dbOtherDis) == false)
                                        {
                                            if (dbOtherDis.IndexOf("*" + varQDC + "*") > -1)
                                                sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"AdditionalDiagnosisPDSS\" value=\"" + varQDC + "\" />");

                                        }

                                    } //Drop down Loop

                                } //Agency Value = 2


                                if ((varAgencyValue == 1 || varAgencyValue == 2 || varAgencyValue == 3) && dbOutletCode2 != "15038")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"LivingArrangements\" value=\"" + dbLivingArrangePres + "\" />");

                                if (varAgencyValue == 2)
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"DependentChildren\" value=\"" + dbDepChild + "\" />");

                                if (dbOutletCode2 == "15034" || dbOutletCode2 == "15038" || dbOutletCode2 == "15055")
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"LivesWithAtEntry\" value=\"" + dbLivingArrangePres + "\" />");

                                    if (String.IsNullOrEmpty(dbEndDate) == false)
                                        sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"LivesWithAtExit\" value=\"" + dbLivesWithAtExit + "\" />");
                                    else
                                        sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"LivesWithAtExit\" value=\"\" />");

                                }

                                if (varAgencyValue == 1 || varAgencyValue == 2 || varAgencyValue == 3)
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ResidentialSetting\" value=\"" + dbResidentialPres + "\" />");

                                    //[Redundant]
                                    //sw.WriteLine("      <service_user_response question_fieldname=\"ResidentialPostcode\" value=\"" + arrPostCode + "\" />");
                                }

                                if (dbOutletCode2 == "15034" || dbOutletCode2 == "15038" || dbOutletCode2 == "15055")
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"AccommodationPrior\" value=\"" + dbResidentialPres + "\" />");

                                    if (String.IsNullOrEmpty(dbEndDate) == false)
                                        sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"AccommodationExit\" value=\"" + dbLivesWithAtExit + "\" />");
                                    else
                                        sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"AccommodationExit\" value=\"\" />");
                                }

                                if (varAgencyValue == 3)
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"EndAccomHACC\" value=\"" + dbLivesWithAtExit + "\" />");

                                if (dbOutletCode2 == "15034")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"HBOSNominationRights\" value=\"" + dbNominationRights + "\" />");

                                if (varAgencyValue == 1 || varAgencyValue == 2)
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"LabourForceStatus\" value=\"" + dbLabourForcePres + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"IncomeSource\" value=\"" + dbIncomePres + "\" />");

                                    if (varAge > 0 && varAge < 16)
                                        sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"CarerAllowance\" value=\"" + dbCarerAllowance + "\" />");
                                }

                                if (varAgencyValue == 2 || varAgencyValue == 3)
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PensionBenefitStatus\" value=\"" + varPensionBenefitStatus + "\" />");

                                if (varAgencyValue == 1)
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ParticipateOutsideNoTransport\" value=\"" + dbOutsideNoTransport + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ParticipateTransport\" value=\"" + dbParticipateTransport + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ParticipateFamily\" value=\"" + dbParticipateFamily + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ParticipateSocial\" value=\"" + dbParticipateSocial + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ParticipateLeisure\" value=\"" + dbParticipateLeisure + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ParticipateWorking\" value=\"" + dbParticipateWorking + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ParticipateMoney\" value=\"" + dbParticipateMoney + "\" />");
                                }

                                if (varAgencyValue == 1 || varAgencyValue == 2)
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"SupportSelfCare\" value=\"" + dbSupportSelfCare + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"SupportMobility\" value=\"" + dbSupportMobility + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"SupportCommunication\" value=\"" + dbSupportCommunication + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"SupportInterpersonal\" value=\"" + dbSupportInterpersonal + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"SupportLearning\" value=\"" + dbSupportLearning + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"SupportEducation\" value=\"" + dbSupportEducation + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"SupportCommunity\" value=\"" + dbSupportCommunity + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"SupportDomestic\" value=\"" + dbSupportDomestic + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"SupportWorking\" value=\"" + dbSupportWorking + "\" />");
                                }

                                if (varAgencyValue == 2)
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"SourceClinicalSupport\" value=\"" + dbClinicalSupport + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ContactCaseManager\" value=\"" + varContactCaseManger + "\" />");
                                }

                                if (varAgencyValue == 1 || varAgencyValue == 2)
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"IndividualFundingStatus\" value=\"" + dbIndividualFunding + "\" />");


                                if (dbOutletCode2 == "17008")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPCreateDateAccOutreach\" value=\"" + varPPCreateDateAccOutreach + "\" />");

                                if (dbOutletCode2 == "17082")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPCreateDateCO\" value=\"" + varPPCreateDateCO + "\" />");

                                if (dbOutletCode2 == "17017")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPCreateDateCongCare\" value=\"" + varPPCreateDateCongCare + "\" />");

                                if (dbOutletCode2 == "17022")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPCreateDateDP\" value=\"" + varPPCreateDateDP + "\" />");

                                if (dbOutletCode2 == "17052")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPCreateDateFamOpt\" value=\"" + varPPCreateDateFamOpt + "\" />");

                                if (dbOutletCode2 == "17201")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPCreateDateFFYA\" value=\"" + varPPCreateDateFFYA + "\" />");

                                if (dbOutletCode2 == "17200")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPCreateDateHomeFirst\" value=\"" + varPPCreateDateHomeFirst + "\" />");

                                if (dbOutletCode2 == "17081")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPCreateDateISP\" value=\"" + varPPCreateDateISP + "\" />");

                                if (dbOutletCode2 == "17083")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPCreateDateMA\" value=\"" + varPPCreateDateMA + "\" />");

                                if (dbOutletCode2 == "15034" || dbOutletCode2 == "15037" || dbOutletCode2 == "15035" || dbOutletCode2 == "15038" || dbOutletCode2 == "15055")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPCreateDatePDSS\" value=\"" + varPPCreateDatePDSS + "\" />");

                                if (dbOutletCode2 == "17016")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPCreateDateSharedSup\" value=\"" + varPPCreateDateSharedSup + "\" />");


                                if (dbOutletCode2 == "17008")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPReviewDateAccOutreach\" value=\"" + varPPReviewDateAccOutreach + "\" />");

                                if (dbOutletCode2 == "17082")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPReviewDateCO\" value=\"" + varPPReviewDateCO + "\" />");

                                if (dbOutletCode2 == "17017")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPReviewDateCongCare\" value=\"" + varPPReviewDateCongCare + "\" />");

                                if (dbOutletCode2 == "17022")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPReviewDateDP\" value=\"" + varPPReviewDateDP + "\" />");

                                if (dbOutletCode2 == "17052")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPReviewDateFamOpt\" value=\"" + varPPReviewDateFamOpt + "\" />");

                                if (dbOutletCode2 == "17201")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPReviewDate-FFYA\" value=\"" + varPPReviewDate_FFYA + "\" />");

                                if (dbOutletCode2 == "17200")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPReviewDateHomeFirst\" value=\"" + varPPReviewDateHomeFirst + "\" />");

                                if (dbOutletCode2 == "17081")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPReviewDateISP\" value=\"" + varPPReviewDateISP + "\" />");

                                if (dbOutletCode2 == "17083")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPReviewDateMA\" value=\"" + varPPReviewDateMA + "\" />");

                                if (dbOutletCode2 == "15034" || dbOutletCode2 == "15037" || dbOutletCode2 == "15035" || dbOutletCode2 == "15038" || dbOutletCode2 == "15055")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPReviewDatePDSS\" value=\"" + varPPReviewDatePDSS + "\" />");

                                if (dbOutletCode2 == "17016")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPReviewDateSharedSup\" value=\"" + varPPReviewDateSharedSup + "\" />");


                                if (varAgencyValue == 2)
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"PPGoals\" value=\"" + varPPGoals + "\" />");

                                if (dbOutletCode2 == "17022")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ClientEFT-DayProg\" value=\"" + dbClientEft + "\" />");


                                if (dbOutletCode2 == "17006")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ReferralDate-DSBCrimJustice\" value=\"" + dbReferralDate + "\" />");

                                if (dbOutletCode2 == "17026")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ReferralDateDSB-BIS\" value=\"" + dbReferralDate + "\" />");

                                if (dbOutletCode2 == "17028")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ReferralDateDSB-CM\" value=\"" + dbReferralDate + "\" />");

                                if (dbOutletCode2 == "15034")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ReferralDatePDSS-HBOS\" value=\"" + dbReferralDate + "\" />");

                                if (dbOutletCode2 == "15037")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ReferralDatePDSS-Respite\" value=\"" + dbReferralDate + "\" />");

                                if (dbOutletCode2 == "15035")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ReferralDatePDSS-PSRDP\" value=\"" + dbReferralDate + "\" />");

                                if (dbOutletCode2 == "15038")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ReferralDatePDSS-ResiReh\" value=\"" + dbReferralDate + "\" />");

                                if (dbOutletCode2 == "15055")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ReferralDatePDSS-SAccom\" value=\"" + dbReferralDate + "\" />");

                                if (dbOutletCode2 == "17042")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ReferralDate-DSBTherapy\" value=\"" + dbReferralDate + "\" />");


                                //If varAgencyValue == 2 Then
                                //[Redundant]
                                //sw.WriteLine("      <service_user_response question_fieldname=\"ReferralSourcePDSS\" value=\"" + arrReferralSourcePDSS + "\" />");

                                if (varAgencyValue == 3)
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ReferralSourceHACC\" value=\"" + dbSourceRef + "\" />");

                                if (dbOutletCode2 == "15037")
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ReasonRespite\" value=\"" + varReasonRespite + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"NightsReceived-PDSSRespite\" value=\"" + varNightsReceived_PDSSRespite + "\" />");
                                }

                                if (dbOutletCode2 == "15038" || dbOutletCode2 == "15055")
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"NightsReceived-PDSSResiRehab\" value=\"" + varNightsReceived_PDSSResiRehab + "\" />");

                                if (varAgencyValue == 3)
                                {
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"ClientReasonHACC\" value=\"" + dbClientStatus + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"LastAssessmentDate\" value=\"" + varLastAssessmentDate + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"NbrMealsAtCentre\" value=\"" + varNbrMealsAtCentre + "\" />");
                                    sbHeaderList.AppendLine("      <service_user_response question_fieldname=\"NbrMealsAtHome\" value=\"" + varNbrMealsAtHome + "\" />");
                                }

                                //If varAgencyValue == 1 OR varAgencyValue == 2 Then
                                //'[Redundant]
                                //'sw.WriteLine("      <service_user_response question_fieldname=\"EndReasonCSTDA\" value=\"" + arrEndReasonCSTDA + "\" />");
                                //End If

                                //If varAgencyValue == 3 Then
                                //'[Redundant]
                                //'sw.WriteLine("      <service_user_response question_fieldname=\"EndReasonHACC\" value=\"" + arrEndReasonHACC + "\" />");
                                //End If


                                //Only append end date and cessation information if valid end date and agency id
                                varDoEnd = "";
                                varDoCessation = "";
                                if (String.IsNullOrEmpty(dbEndDate) == false)
                                {
                                    varDoEnd = " end_date=\"" + dbEndDate + "\"";

                                    if (varAgencyValue == 1 || varAgencyValue == 2)
                                        varDoCessation = " end_reason_code = \"" + dbCessation + "\"";
                                    else
                                        varDoCessation = " end_reason_code = \"" + dbCessation + "\"";
                                }

                                //Append referral information
                                varDoReferral = "";
                                if (varAgencyValue == 2)
                                    varDoReferral = " referral_source_code = \"" + dbSourceRef + "\"";
                                else if (varAgencyValue == 3)
                                    varDoReferral = " referral_source_code = \"" + dbSourceRef + "\"";

                                //Only append snapshot information for final quarter
                                varDoSnapShot = "";
                                if (varQuarter == 4)
                                    varDoSnapShot = " snapshot_type=\"" + varSnapshot_Type + "\"";


                                sbHeaderList.AppendLine("      <service_user_outlet start_date=\"" + dbStartDate + "\"" + varDoEnd + " last_service_date=\"" + varLast_Service_Date + "\"" + varDoCessation + varDoSnapShot + varDoReferral + " service_type_outlet_codeno=\"" + varService_Type_Outlet_Codeno + "\">");


                                //3/
                                //Loop through activities
                                //Get this from activities table

                                //Get the start and end date for each month. Also reset the variables
                                varM1Start = varStartDate;
                                varM1End = varM1Start.AddMonths(1);
                                varM1End = varM1End.AddDays(-1);

                                varM1Hours = 0;
                                varM1First = "";
                                varM1Last = "";

                                varM2Start = varM1Start.AddMonths(1);
                                varM2End = varM2Start.AddMonths(1);
                                varM2End = varM2End.AddDays(-1);

                                varM2Hours = 0;
                                varM2First = "";
                                varM2Last = "";

                                varM3Start = varM2Start.AddMonths(1);
                                varM3End = varM3Start.AddMonths(1);
                                varM3End = varM3End.AddDays(-1);

                                varM3Hours = 0;
                                varM3First = "";
                                varM3Last = "";


                                //Loop through activities to get start date
                                foreach (var p in result8.Entities)
                                {
                                    // We need to get the entity id for the client field for comparisons
                                    if (p.Attributes.Contains("new_supportperiod"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)p.Attributes["new_supportperiod"];
                                        dbSupportPeriod = getEntity.Id.ToString();
                                    }
                                    else if (p.FormattedValues.Contains("new_supportperiod"))
                                        dbSupportPeriod = p.FormattedValues["new_supportperiod"];
                                    else
                                        dbSupportPeriod = "";

                                    if (p.FormattedValues.Contains("new_entrydate"))
                                        dbEntryDate = p.FormattedValues["new_entrydate"];
                                    else if (c.Attributes.Contains("new_entrydate"))
                                        dbEntryDate = p.Attributes["new_entrydate"].ToString();
                                    else
                                        dbEntryDate = "";

                                    // Convert date from American format to Australian format
                                    dbEntryDate = cleanDateAM(dbEntryDate);

                                    if (p.FormattedValues.Contains("new_amount"))
                                        dbAmount = p.FormattedValues["new_amount"];
                                    else if (c.Attributes.Contains("new_amount"))
                                        dbAmount = p.Attributes["new_amount"].ToString();
                                    else
                                        dbAmount = "";

                                    dbAmount = cleanString(dbAmount, "double");
                                    Double.TryParse(dbAmount, out varCheckDouble);

                                    // Only get if support periods match, date is valid and amount is > 0
                                    if (dbSupportPeriod == dbPalmClientSupportId && string.IsNullOrEmpty(dbEntryDate) == false && varCheckDouble > 0)
                                    {
                                        // Add the total to the correct period
                                        if (Convert.ToDateTime(dbEntryDate) >= Convert.ToDateTime(varM1Start) && Convert.ToDateTime(dbEntryDate) <= Convert.ToDateTime(varM1End))
                                        {
                                            varM1Hours += varCheckDouble;

                                            if (String.IsNullOrEmpty(varM1First) == true)
                                                varM1First = dbEntryDate;

                                            if (String.IsNullOrEmpty(varM1Last) == true)
                                                varM1Last = dbEntryDate;
                                            else if (Convert.ToDateTime(dbEntryDate) > Convert.ToDateTime(varM1Last))
                                                varM1Last = dbEntryDate;

                                        }
                                        else if (Convert.ToDateTime(dbEntryDate) >= Convert.ToDateTime(varM2Start) && Convert.ToDateTime(dbEntryDate) <= Convert.ToDateTime(varM2End))
                                        {
                                            varM2Hours += varCheckDouble;

                                            if (String.IsNullOrEmpty(varM2First) == true)
                                                varM2First = dbEntryDate;

                                            if (String.IsNullOrEmpty(varM2Last) == true)
                                                varM2Last = dbEntryDate;
                                            else if (Convert.ToDateTime(dbEntryDate) > Convert.ToDateTime(varM2Last))
                                                varM2Last = dbEntryDate;

                                        }
                                        else if (Convert.ToDateTime(dbEntryDate) >= Convert.ToDateTime(varM3Start) && Convert.ToDateTime(dbEntryDate) <= Convert.ToDateTime(varM3End))
                                        {
                                            varM3Hours += varCheckDouble;

                                            if (String.IsNullOrEmpty(varM3First) == true)
                                                varM3First = dbEntryDate;

                                            if (String.IsNullOrEmpty(varM3Last) == true)
                                                varM3Last = dbEntryDate;
                                            else if (Convert.ToDateTime(dbEntryDate) > Convert.ToDateTime(varM3Last))
                                                varM3Last = dbEntryDate;

                                        }

                                    } //From Date not null

                                } //Activities List Loop


                                //The hours should actually be minutes
                                varM1Hours = varM1Hours * 60;
                                varM2Hours = varM2Hours * 60;
                                varM3Hours = varM3Hours * 60;

                                //Have as integer to remove decimal places
                                varM1Hours = (int)varM1Hours;
                                varM2Hours = (int)varM2Hours;
                                varM3Hours = (int)varM3Hours;

                                //Append service hours information
                                if (String.IsNullOrEmpty(varM1First) == false)
                                    sbHeaderList.AppendLine("         <service_hours from_date=\"" + varM1First + "\" to_date=\"" + varM1Last + "\" service_hours=\"" + varM1Hours + "\" />");

                                if (String.IsNullOrEmpty(varM2First) == false)
                                    sbHeaderList.AppendLine("         <service_hours from_date=\"" + varM2First + "\" to_date=\"" + varM2Last + "\" service_hours=\"" + varM2Hours + "\" />");

                                if (String.IsNullOrEmpty(varM3First) == false)
                                    sbHeaderList.AppendLine("         <service_hours from_date=\"" + varM3First + "\" to_date=\"" + varM3Last + "\" service_hours=\"" + varM3Hours + "\" />");

                                sbHeaderList.AppendLine("      </service_user_outlet>");

                                sbHeaderList.AppendLine("   </service_user>");

                            } // varDoNext

                        } // Client Loop


                        //Close off extract
                        sbHeaderList.AppendLine("</agency>");

                        //varTest = sbHeaderList.ToString();

                        // Create note against current Palm Go QDC record and add attachment
                        byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                        string encodedData = System.Convert.ToBase64String(filename);
                        Entity Annotation = new Entity("annotation");
                        Annotation.Attributes["objectid"] = new EntityReference("new_palmgoqdc", varQdcID);
                        Annotation.Attributes["objecttypecode"] = "new_palmgoqdc";
                        Annotation.Attributes["subject"] = "QDC Extract";
                        Annotation.Attributes["documentbody"] = encodedData;
                        Annotation.Attributes["mimetype"] = @"text / plain";
                        Annotation.Attributes["notetext"] = "QDC Extract for " + varYear + " quarter " + varQuarter;
                        Annotation.Attributes["filename"] = varFileName;
                        _service.Create(Annotation);

                        // If there is an error, create note against current Palm Go QDC record and add attachment
                        if (sbErrorList.Length > 0)
                        {
                            byte[] filename2 = Encoding.ASCII.GetBytes(sbErrorList.ToString());
                            string encodedData2 = System.Convert.ToBase64String(filename2);
                            Entity Annotation2 = new Entity("annotation");
                            Annotation2.Attributes["objectid"] = new EntityReference("new_palmgoqdc", varQdcID);
                            Annotation2.Attributes["objecttypecode"] = "new_palmgoqdc";
                            Annotation2.Attributes["subject"] = "QDC Extract";
                            Annotation2.Attributes["documentbody"] = encodedData2;
                            Annotation2.Attributes["mimetype"] = @"text / plain";
                            Annotation2.Attributes["notetext"] = "QDC errors and warnings for " + varYear + " quarter " + varQuarter;
                            Annotation2.Attributes["filename"] = varFileName2;
                            _service.Create(Annotation2);
                        }

                        // throw new InvalidPluginExecutionException("This plugin is working\r\n" + varTest);
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
    }
}

