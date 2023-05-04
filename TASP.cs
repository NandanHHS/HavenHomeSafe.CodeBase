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
	public class goTASP: IPlugin
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
                // if (entity.LogicalName != "new_palmgotasp")
                //    return;

                try
                {
                    // Fixes:
                    // Remove ampersand from comments
                    // Compare values to Lower

                    string varDescription = ""; // Description field on form
                    bool varCreateExtract = false; // Create Extract field on form
                    Guid varTaspID = new Guid(); // GUID for palm go tasp record
                    StringBuilder sbHeaderList = new StringBuilder(); // String builder for report header
                    StringBuilder sbReportList = new StringBuilder(); // String builder for report data
                    StringBuilder sbErrorList = new StringBuilder(); // String builder for error list

                    string varFileName = ""; // Extract file name
                    string varFileName2 = ""; // Error log file name
                    DateTime varStartDate = new DateTime(); // Start date of extract
                    DateTime varStartDatePr = new DateTime(); // Print start date of extract
                    DateTime varEndDate = new DateTime(); // End date of extract
                    DateTime varEndDatePr = new DateTime(); // Print end date of extract
                    int varCheckInt = 0; // Parse integers
                    double varCheckDouble = 0; // Parse doubles
                    DateTime varCheckDate = new DateTime(); // Parse dates
                    EntityReference getEntity; // Entity reference object
                    AliasedValue getAlias; // Aliased value object

                    string varTest = ""; // Debug

                    // Only do this if the entity is the Palm Go TASP entity
                    if (entity.LogicalName == "new_palmgotasp")
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

                        // Get info for current TASP record
                        varDescription = entity.GetAttributeValue<string>("new_description");
                        varCreateExtract = entity.GetAttributeValue<bool>("new_createextract");

                        // Important: The plugin uses American dates but returns formatted Australian dates
                        // Any dates created in the plugin will be American

                        varStartDate = entity.GetAttributeValue<DateTime>("new_datefrom");
                        varStartDatePr = Convert.ToDateTime(varStartDate.AddHours(14).ToString()); // Australian Date

                        varEndDate = entity.GetAttributeValue<DateTime>("new_dateto");
                        varEndDatePr = Convert.ToDateTime(varEndDate.AddHours(23).ToString()); // Australian Date

                        varEndDate = varEndDate.AddHours(23); // Correct for Australian time

                        // Get GUID
                        varTaspID = entity.Id;

                        //EntityReference ownerLookup = (EntityReference)entity.Attributes["new_agencyid"];
                        //varAgencyId += ownerLookup.Id.ToString() + ".\r\n";
                        //varAgencyId += ((EntityReference)entity.Attributes["new_agencyid"]).Name + ".\r\n";
                        //varAgencyId += ownerLookup.LogicalName + ".\r\n";

                        //var actualOwningUnit = _service.Retrieve(ownerLookup.LogicalName, ownerLookup.Id, new ColumnSet(true));

                        //varAgencyId = actualOwningUnit["new_agency"].ToString();

                        // Create file name for report and error log
                        varFileName = "TAAP Report " + varStartDatePr.ToString("MMM") + " " + varStartDatePr.Year + "-" + varEndDatePr.ToString("MMM") + " " + varEndDatePr.Year + ".xls";

                        varFileName2 = varFileName.Replace(".xls", ".txt");
                        varFileName2 = "Errors for " + varFileName2;

                        // Fetch statements for database
                        // Get the required fields from the client table (and associated entities)
                        // Any clients that have a support period ticked as ARC for the period
                        string dbClientList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclient'>
                                <attribute name='new_palmclientid' />
                                <attribute name='new_address' />
                                <attribute name='new_indigenous' />
                                <attribute name='new_gender' />
                                <attribute name='new_dob' />
                                <link-entity name='new_palmclientsupport' to='new_palmclientid' from='new_client' link-type='inner'>
                                    <attribute name='new_arc' />
                                    <attribute name='new_agrefno' />
                                    <attribute name='new_startdate' />
                                    <attribute name='new_servtype' />
                                    <attribute name='new_locality' />
                                    <attribute name='new_taspcald' />
                                    <attribute name='new_homestatus' />
                                    <attribute name='new_taspref' />
                                    <attribute name='new_tenancymat' />
                                    <attribute name='new_findis' />
                                    <attribute name='new_wellbeing' />
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_enddate' />
                                    <attribute name='new_vcathearing' />
                                    <link-entity name='new_palmddllocality' to='new_locality' from='new_palmddllocalityid' link-type='outer'>
                                        <attribute name='new_postcode' />
                                        <attribute name='new_state' />
                                    </link-entity>
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientsupport' attribute='new_arc' operator='eq' value='True' />
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get the required fields from the activities table (and associated entities)
                        // Any activities for the period that have a support period ticked
                        string dbActivityList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                                <entity name='new_palmclientactivities'>
                                    <attribute name='new_supportperiod' />
                                    <attribute name='new_amount' />
                                    <attribute name='new_activityarc' />
                                    <attribute name='new_entrydate' />
                                    <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                            <attribute name='new_client' />
                                    </link-entity>
                                    <filter type='and'>
                                        <condition entityname='new_palmclientactivities' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                        <condition entityname='new_palmclientactivities' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                        <condition entityname='new_palmclientactivities' attribute='new_activityarc' operator='ge' value='0' />
                                        <condition entityname='new_palmclientsupport' attribute='new_arc' operator='eq' value='True' />
                                    </filter>
                                </entity>
                            </fetch> ";

                        // Get the required fields from the VCAT table (and associated entities)
                        // Any activities for the period
                        string dbVCATList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                                <entity name='new_palmclientvcat'>
                                    <attribute name='new_client' />
                                    <attribute name='new_vcatvenue' />
                                    <attribute name='new_outcome' />
                                    <attribute name='new_entrydate' />
                                    <link-entity name='new_palmclient' to='new_client' from='new_palmclientid' link-type='inner'>
                                            <attribute name='new_palmclientid' />
                                    </link-entity>
                                    <filter type='and'>
                                        <condition entityname='new_palmclientvcat' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                        <condition entityname='new_palmclientvcat' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    </filter>
                                </entity>
                            </fetch> ";

                        // Get the required fields from the travel table (and associated entities)
                        // Any travel for the period that have a support period ticked
                        string dbTravelList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                                <entity name='new_palmclienttravel'>
                                    <attribute name='new_supportperiod' />
                                    <attribute name='new_distance' />
                                    <attribute name='new_tvalue' />
                                    <attribute name='new_entrydate' />
                                    <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                            <attribute name='new_client' />
                                    </link-entity>
                                    <filter type='and'>
                                        <condition entityname='new_palmclienttravel' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                        <condition entityname='new_palmclienttravel' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    </filter>
                                </entity>
                            </fetch> ";

                        // Variables used to store the data from the fetchXML statements
                        string dbPalmClientId = "";
                        string dbClient = "";
                        string dbIndigenous = "";
                        string dbGender = "";
                        string dbDob = "";
                        string dbArc = "";
                        string dbAgRefNo = "";
                        string dbStartDate = "";
                        string dbServType = "";
                        string dbLocality = "";
                        string dbTaspCald = "";
                        string dbHomeStatus = "";
                        string dbTaspRef = "";
                        string dbTenancyMat = "";
                        string dbFinDis = "";
                        string dbWellBeing = "";
                        string dbPalmClientSupportId = "";
                        string dbEndDate = "";
                        string dbVcatHearing = "";

                        string dbPostcode = "";

                        string dbActSupportPeriod = "";
                        string dbAmount = "";
                        string dbActivityArc = "";
                        string dbEntryDate = "";

                        string dbVcatClient = "";
                        string dbVcatVenue = "";
                        string dbOutcome = "";

                        string dbTravSupportPeriod = "";
                        string dbDistance = "";
                        string dbTValue = "";

                        // Counters for the information above
                        int varGetAge = 0;
                        double varAmount = 0;
                        string varAge = "";
                        string varHDate = "";
                        string varHVenue = "";
                        string varOutcome = "";

                        int varInfo = 0;
                        int varNego = 0;
                        int varVCATPrep = 0;
                        int varVCATRep = 0;
                        int varTravel = 0;
                        int varAdmin = 0;

                        int varDistance = 0;
                        int varTValue = 0;

                        // Get the fetch XML data and place in entity collection objects
                        EntityCollection result = _service.RetrieveMultiple(new FetchExpression(dbClientList));
                        EntityCollection result2 = _service.RetrieveMultiple(new FetchExpression(dbActivityList));
                        EntityCollection result3 = _service.RetrieveMultiple(new FetchExpression(dbVCATList));
                        EntityCollection result4 = _service.RetrieveMultiple(new FetchExpression(dbTravelList));

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

                            //Reset variables
                            varGetAge = 0;
                            varAmount = 0;
                            varAge = "";
                            varHDate = "";
                            varHVenue = "";
                            varOutcome = "";

                            varInfo = 0;
                            varNego = 0;
                            varVCATPrep = 0;
                            varVCATRep = 0;
                            varTravel = 0;
                            varAdmin = 0;

                            varDistance = 0;
                            varTValue = 0;

                            if (c.FormattedValues.Contains("new_address"))
                                dbClient = c.FormattedValues["new_address"];
                            else if (c.Attributes.Contains("new_address"))
                                dbClient = c.Attributes["new_address"].ToString();
                            else
                                dbClient = "";

                            if (c.FormattedValues.Contains("new_indigenous"))
                                dbIndigenous = c.FormattedValues["new_indigenous"];
                            else if (c.Attributes.Contains("new_indigenous"))
                                dbIndigenous = c.Attributes["new_indigenous"].ToString();
                            else
                                dbIndigenous = "";

                            // TASP values for indigenous
                            if (dbIndigenous.ToLower() == "aboriginal but not torres strait islander")
                                dbIndigenous = "Aboriginal";
                            if (dbIndigenous.ToLower() == "torres strait islander but not aboriginal")
                                dbIndigenous = "Torres Strait Islander";
                            if (dbIndigenous.ToLower() == "not aboriginal or torres strait islander")
                                dbIndigenous = "Neither";
                            if (dbIndigenous.ToLower() == "both aboriginal and torres strait islander")
                                dbIndigenous = "Both";

                            if (c.FormattedValues.Contains("new_gender"))
                                dbGender = c.FormattedValues["new_gender"];
                            else if (c.Attributes.Contains("new_gender"))
                                dbGender = c.Attributes["new_gender"].ToString();
                            else
                                dbGender = "";

                            if (c.FormattedValues.Contains("new_dob"))
                                dbDob = c.FormattedValues["new_dob"];
                            else if (c.Attributes.Contains("new_dob"))
                                dbDob = c.Attributes["new_dob"].ToString();
                            else
                                dbDob = "";

                            // Convert date from American format to Australian format
                            dbDob = cleanDateAM(dbDob);

                            // get age
                            if (String.IsNullOrEmpty(dbDob) == false)
                            {
                                DateTime.TryParse(dbDob, out varCheckDate);
                                varGetAge = DateTime.Now.Year - varCheckDate.Year;
                                if (DateTime.Now.Month < varCheckDate.Month || (DateTime.Now.Month == varCheckDate.Month && DateTime.Now.Day < varCheckDate.Day))
                                    varGetAge = varGetAge - 1;

                                if (varGetAge > -1)
                                    varAge = varGetAge + "";
                            }

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_arc"))
                                dbArc = c.FormattedValues["new_palmclientsupport1.new_arc"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_arc"))
                                dbArc = c.Attributes["new_palmclientsupport1.new_arc"].ToString();
                            else
                                dbArc = "";

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_agrefno"))
                                dbAgRefNo = c.FormattedValues["new_palmclientsupport1.new_agrefno"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_agrefno"))
                                dbAgRefNo = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_agrefno").Value.ToString();
                            else
                                dbAgRefNo = "";

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_startdate"))
                                dbStartDate = c.FormattedValues["new_palmclientsupport1.new_startdate"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_startdate"))
                                dbStartDate = c.Attributes["new_palmclientsupport1.new_startdate"].ToString();
                            else
                                dbStartDate = "";

                            // Convert date from American format to Australian format
                            dbStartDate = cleanDateAM(dbStartDate);

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_servtype"))
                                dbServType = c.FormattedValues["new_palmclientsupport1.new_servtype"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_servtype"))
                                dbServType = c.Attributes["new_palmclientsupport1.new_servtype"].ToString();
                            else
                                dbServType = "";

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_locality"))
                                dbLocality = c.FormattedValues["new_palmclientsupport1.new_locality"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_locality"))
                                dbLocality = c.Attributes["new_palmclientsupport1.new_locality"].ToString();
                            else
                                dbLocality = "";

                            if (c.FormattedValues.Contains("new_palmddllocality2.new_postcode"))
                                dbPostcode = c.FormattedValues["new_palmddllocality2.new_postcode"];
                            else if (c.Attributes.Contains("new_palmddllocality2.new_postcode"))
                                dbPostcode = c.Attributes["new_palmddllocality2.new_postcode"].ToString();
                            else
                                dbPostcode = "";

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_taspcald"))
                                dbTaspCald = c.FormattedValues["new_palmclientsupport1.new_taspcald"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_taspcald"))
                                dbTaspCald = c.Attributes["new_palmclientsupport1.new_taspcald"].ToString();
                            else
                                dbTaspCald = "";

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_homestatus"))
                                dbHomeStatus = c.FormattedValues["new_palmclientsupport1.new_homestatus"];
                            else if (c.Attributes.Contains("new_supponew_palmclientsupport1rtperiod1.new_homestatus"))
                                dbHomeStatus = c.Attributes["new_palmclientsupport1.new_homestatus"].ToString();
                            else
                                dbHomeStatus = "";

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_taspref"))
                                dbTaspRef = c.FormattedValues["new_palmclientsupport1.new_taspref"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_taspref"))
                                dbTaspRef = c.Attributes["new_palmclientsupport1.new_taspref"].ToString();
                            else
                                dbTaspRef = "";

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_tenancymat"))
                                dbTenancyMat = c.FormattedValues["new_palmclientsupport1.new_tenancymat"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_tenancymat"))
                                dbTenancyMat = c.Attributes["new_palmclientsupport1.new_tenancymat"].ToString();
                            else
                                dbTenancyMat = "";

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_findis"))
                                dbFinDis = c.FormattedValues["new_palmclientsupport1.new_findis"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_findis"))
                                dbFinDis = c.Attributes["new_palmclientsupport1.new_findis"].ToString();
                            else
                                dbFinDis = "";

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_wellbeing"))
                                dbWellBeing = c.FormattedValues["new_palmclientsupport1.new_wellbeing"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_wellbeing"))
                                dbWellBeing = c.Attributes["new_palmclientsupport1.new_wellbeing"].ToString();
                            else
                                dbWellBeing = "";

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_palmclientsupportid"))
                                dbPalmClientSupportId = c.FormattedValues["new_palmclientsupport1.new_palmclientsupportid"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_palmclientsupportid"))
                                dbPalmClientSupportId = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_palmclientsupportid").Value.ToString();
                            else
                                dbPalmClientSupportId = "";

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_enddate"))
                                dbEndDate = c.FormattedValues["new_palmclientsupport1.new_enddate"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_enddate"))
                                dbEndDate = c.Attributes["new_palmclientsupport1.new_enddate"].ToString();
                            else
                                dbEndDate = "";

                            // Convert date from American format to Australian format
                            dbEndDate = cleanDateAM(dbEndDate);

                            if (c.FormattedValues.Contains("new_palmclientsupport1.new_vcathearing"))
                                dbVcatHearing = c.FormattedValues["new_palmclientsupport1.new_vcathearing"];
                            else if (c.Attributes.Contains("new_palmclientsupport1.new_vcathearing"))
                                dbVcatHearing = c.Attributes["new_palmclientsupport1.new_vcathearing"].ToString();
                            else
                                dbVcatHearing = "";

                            //varTest += dbPalmClientSupportId + "\r\n";

                            // Loop through activities data
                            foreach (var a in result2.Entities)
                            {
                                // We need to get the entity id for the support period field for comparisons
                                if (a.Attributes.Contains("new_supportperiod"))
                                {
                                    // Get the entity id for the client using the entity reference object
                                    getEntity = (EntityReference)a.Attributes["new_supportperiod"];
                                    dbActSupportPeriod = getEntity.Id.ToString();
                                }
                                else if (a.FormattedValues.Contains("new_supportperiod"))
                                    dbActSupportPeriod = a.FormattedValues["new_supportperiod"];
                                else
                                    dbActSupportPeriod = "";

                                // Need to see if same support period
                                if (dbPalmClientSupportId == dbActSupportPeriod)
                                {
                                    // Process the data as follows:
                                    // If there is a formatted value for the field, use it
                                    // Otherwise if there is a literal value for the field, use it
                                    // Otherwise the value wasn't returned so set as nothing
                                    if (a.FormattedValues.Contains("new_amount"))
                                        dbAmount = a.FormattedValues["new_amount"];
                                    else if (a.Attributes.Contains("new_amount"))
                                        dbAmount = a.Attributes["new_amount"].ToString();
                                    else
                                        dbAmount = "";

                                    // Ensure numeric and express as hour
                                    dbAmount = cleanString(dbAmount, "double");

                                    Double.TryParse(dbAmount, out varAmount);
                                    if (varAmount > 0)
                                        varAmount = 60 * varAmount;

                                    if (a.FormattedValues.Contains("new_activityarc"))
                                        dbActivityArc = a.FormattedValues["new_activityarc"];
                                    else if (a.Attributes.Contains("new_activityarc"))
                                        dbActivityArc = a.Attributes["new_activityarc"].ToString();
                                    else
                                        dbActivityArc = "";

                                    // Add to totals
                                    if (dbActivityArc == "Information and Advice")
                                        varInfo += (int)varAmount;
                                    if (dbActivityArc == "Negotiation")
                                        varNego += (int)varAmount;
                                    if (dbActivityArc == "VCAT Preparation")
                                        varVCATPrep += (int)varAmount;
                                    if (dbActivityArc == "VCAT Representation")
                                        varVCATRep += (int)varAmount;
                                    if (dbActivityArc == "Travel")
                                        varTravel += (int)varAmount;
                                    if (dbActivityArc == "Administration")
                                        varAdmin += (int)varAmount;

                                    if (a.FormattedValues.Contains("new_entrydate"))
                                        dbEntryDate = a.FormattedValues["new_entrydate"];
                                    else if (a.Attributes.Contains("new_entrydate"))
                                        dbEntryDate = a.Attributes["new_entrydate"].ToString();
                                    else
                                        dbEntryDate = "";

                                    // Convert date from American format to Australian format
                                    dbEntryDate = cleanDateAM(dbEntryDate);

                                } // Same support period

                            } // Activities Loop


                            // Loop through VCAT data
                            foreach (var v in result3.Entities)
                            {
                                // We need to get the entity id for the client field for comparisons
                                if (v.Attributes.Contains("new_client"))
                                {
                                    // Get the entity id for the client using the entity reference object
                                    getEntity = (EntityReference)v.Attributes["new_client"];
                                    dbVcatClient = getEntity.Id.ToString();
                                }
                                else if (v.FormattedValues.Contains("new_client"))
                                    dbVcatClient = v.FormattedValues["new_client"];
                                else
                                    dbVcatClient = "";

                                // Need to see if same client
                                if (dbPalmClientId == dbVcatClient)
                                {
                                    //varTest += dbStartDate + "\r\n";

                                    // Process the data as follows:
                                    // If there is a formatted value for the field, use it
                                    // Otherwise if there is a literal value for the field, use it
                                    // Otherwise the value wasn't returned so set as nothing
                                    if (v.FormattedValues.Contains("new_vcatvenue"))
                                        dbVcatVenue = v.FormattedValues["new_vcatvenue"];
                                    else if (v.Attributes.Contains("new_vcatvenue"))
                                        dbVcatVenue = v.Attributes["new_vcatvenue"].ToString();
                                    else
                                        dbVcatVenue = "";

                                    if (v.FormattedValues.Contains("new_outcome"))
                                        dbOutcome = v.FormattedValues["new_outcome"];
                                    else if (v.Attributes.Contains("new_outcome"))
                                        dbOutcome = v.Attributes["new_outcome"].ToString();
                                    else
                                        dbOutcome = "";

                                    if (v.FormattedValues.Contains("new_entrydate"))
                                        dbEntryDate = v.FormattedValues["new_entrydate"];
                                    else if (v.Attributes.Contains("new_entrydate"))
                                        dbEntryDate = v.Attributes["new_entrydate"].ToString();
                                    else
                                        dbEntryDate = "";

                                    // Convert date from American format to Australian format
                                    dbEntryDate = cleanDateAM(dbEntryDate);

                                    // Get Venue details for latest hearing
                                    if (String.IsNullOrEmpty(dbEntryDate) == false)
                                    {
                                        if (String.IsNullOrEmpty(varHDate) == true)
                                        {
                                            varHDate = dbEntryDate;
                                            varHVenue = dbVcatVenue;
                                            varOutcome = dbOutcome;
                                        }
                                        else if (Convert.ToDateTime(varHDate) > Convert.ToDateTime(dbEntryDate))
                                        {
                                            varHDate = dbEntryDate;
                                            varHVenue = dbVcatVenue;
                                            varOutcome = dbOutcome;
                                        }
                                    }

                                } // Same client

                            } // VCAT Loop


                            // Loop through travel data
                            foreach (var t in result4.Entities)
                            {
                                // We need to get the entity id for the client field for comparisons
                                if (t.Attributes.Contains("new_supportperiod"))
                                {
                                    // Get the entity id for the client using the entity reference object
                                    getEntity = (EntityReference)t.Attributes["new_supportperiod"];
                                    dbTravSupportPeriod = getEntity.Id.ToString();
                                }
                                else if (t.FormattedValues.Contains("new_supportperiod"))
                                    dbTravSupportPeriod = t.FormattedValues["new_supportperiod"];
                                else
                                    dbTravSupportPeriod = "";

                                // Need to see if same support period
                                if (dbPalmClientSupportId == dbTravSupportPeriod)
                                {
                                    // Process the data as follows:
                                    // If there is a formatted value for the field, use it
                                    // Otherwise if there is a literal value for the field, use it
                                    // Otherwise the value wasn't returned so set as nothing
                                    if (t.FormattedValues.Contains("new_distance"))
                                        dbDistance = t.FormattedValues["new_distance"];
                                    else if (t.Attributes.Contains("new_distance"))
                                        dbDistance = t.Attributes["new_distance"].ToString();
                                    else
                                        dbDistance = "";

                                    // Ensure numeric
                                    dbDistance = cleanString(dbDistance, "double");

                                    if (t.FormattedValues.Contains("new_tvalue"))
                                        dbTValue = t.FormattedValues["new_tvalue"];
                                    else if (t.Attributes.Contains("new_tvalue"))
                                        dbTValue = t.Attributes["new_tvalue"].ToString();
                                    else
                                        dbTValue = "";

                                    // Ensure numeric
                                    dbTValue = cleanString(dbTValue, "double");

                                    if (t.FormattedValues.Contains("new_entrydate"))
                                        dbEntryDate = t.FormattedValues["new_entrydate"];
                                    else if (t.Attributes.Contains("new_entrydate"))
                                        dbEntryDate = t.Attributes["new_entrydate"].ToString();
                                    else
                                        dbEntryDate = "";

                                    // Convert date from American format to Australian format
                                    dbEntryDate = cleanDateAM(dbEntryDate);

                                    // Add to totals
                                    Double.TryParse(dbDistance, out varCheckDouble);
                                    varDistance += (int)varCheckDouble;

                                    Double.TryParse(dbTValue, out varCheckDouble);
                                    varTValue += (int)varCheckDouble;

                                } // Same Support Period

                            } // Travel Loop

                            // Add line to report
                            sbReportList.AppendLine("<tr>\r\n<td>" + dbAgRefNo + "</td>\r\n<td>" + dbStartDate + "</td>\r\n<td>" + dbServType + "</td>\r\n<td>" + dbPostcode + "</td>\r\n<td>" + dbTaspCald + "</td>\r\n<td>" + dbIndigenous + "</td>\r\n<td>" + dbGender + "</td>\r\n<td>" + varAge + "</td>\r\n<td>" + dbHomeStatus + "</td>\r\n<td>" + dbTaspRef + "</td>\r\n<td>" + dbTenancyMat + "</td>\r\n<td>" + dbFinDis + "</td>\r\n<td>" + dbWellBeing + "</td>\r\n<td>" + varInfo + "</td>\r\n<td>" + varNego + "</td>\r\n<td>" + varVCATPrep + "</td>\r\n<td>" + varVCATRep + "</td>\r\n<td>" + varTravel + "</td>\r\n<td>" + varAdmin + "</td>\r\n<td>" + dbEndDate + "</td>\r\n<td>" + varHVenue + "</td>\r\n<td>" + varOutcome + "</td>\r\n<td>" + dbVcatHearing + "</td>\r\n<td>" + varDistance + "</td>\r\n<td>" + varTValue + "</td>\r\n</tr>");

                        } // client loop


                        //Header part of the TASP extract
                        sbHeaderList.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                        sbHeaderList.AppendLine("<head>");
                        sbHeaderList.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                        sbHeaderList.AppendLine("<!--[if gte mso 9]>");
                        sbHeaderList.AppendLine("<xml>");
                        sbHeaderList.AppendLine("<x:ExcelWorkbook>");
                        sbHeaderList.AppendLine("<x:ExcelWorksheets>");
                        sbHeaderList.AppendLine("<x:ExcelWorksheet>");

                        //this line names the worksheet
                        sbHeaderList.AppendLine("<x:Name>TAAP Data</x:Name>");

                        sbHeaderList.AppendLine("<x:WorksheetOptions>");

                        sbHeaderList.AppendLine("<x:Panes>");
                        sbHeaderList.AppendLine("</x:Panes>");
                        sbHeaderList.AppendLine("</x:WorksheetOptions>");
                        sbHeaderList.AppendLine("</x:ExcelWorksheet>");
                        sbHeaderList.AppendLine("</x:ExcelWorksheets>");
                        sbHeaderList.AppendLine("</x:ExcelWorkbook>");
                        sbHeaderList.AppendLine("</xml>");
                        sbHeaderList.AppendLine("<![endif]-->");
                        sbHeaderList.AppendLine("</head>");

                        sbHeaderList.AppendLine("<table width=\"100%\" border=0 cellpadding=5 class=\"myClass1\">");
                        sbHeaderList.AppendLine("<tr>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Client Reference Number</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Date Case Opened</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Number of Clients</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Client Postcode</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">CALD</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">ATSI</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Gender/s</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Age</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Accommodation Type</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Referral Source</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Tenancy Matter</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Financial Disadvantage</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Vulnerability</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Service whole minutes-Info and Advice</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Service whole minutes-Negotiation</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Service whole minutes-VCAT Preparation</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Service whole minutes-VCAT Representation</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Service whole minutes-Travel</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Service whole minutes-Administration</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Date Case Closed</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Current Status</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">VCAT Venue</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">VCAT Hearings-Number</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Travel-Kilometres</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Travel-Parking/Public Transport Costs</td>");
                        sbHeaderList.AppendLine("</tr>");

                        sbHeaderList.AppendLine(sbReportList.ToString());

                        sbHeaderList.AppendLine("</table>");

                        //varTest += sbHeaderList.ToString();

                        // Create note against current Palm Go TASP record and add attachment
                        byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                        string encodedData = System.Convert.ToBase64String(filename);
                        Entity Annotation = new Entity("annotation");
                        Annotation.Attributes["objectid"] = new EntityReference("new_palmgotasp", varTaspID);
                        Annotation.Attributes["objecttypecode"] = "new_palmgotasp";
                        Annotation.Attributes["subject"] = "TASP Extract";
                        Annotation.Attributes["documentbody"] = encodedData;
                        Annotation.Attributes["mimetype"] = @"application/msexcel";
                        Annotation.Attributes["notetext"] = "TASP Extract for " + cleanDate(varStartDatePr) + " to " + cleanDate(varEndDatePr);
                        Annotation.Attributes["filename"] = varFileName;
                        _service.Create(Annotation);

                        // If there is an error, create note against current Palm Go TASP record and add attachment
                        if (sbErrorList.Length > 0)
                        {
                            byte[] filename2 = Encoding.ASCII.GetBytes(sbErrorList.ToString());
                            string encodedData2 = System.Convert.ToBase64String(filename2);
                            Entity Annotation2 = new Entity("annotation");
                            Annotation2.Attributes["objectid"] = new EntityReference("new_palmgotasp", varTaspID);
                            Annotation2.Attributes["objecttypecode"] = "new_palmgotasp";
                            Annotation2.Attributes["subject"] = "TASP Extract";
                            Annotation2.Attributes["documentbody"] = encodedData2;
                            Annotation2.Attributes["mimetype"] = @"text / plain";
                            Annotation2.Attributes["notetext"] = "TASP errors and warnings for " + cleanDate(varStartDatePr) + " to " + cleanDate(varEndDatePr);
                            Annotation2.Attributes["filename"] = varFileName2;
                            _service.Create(Annotation2);
                        }

                        //varTest += varStartDate + " " + varEndDate + "\r\n";
                        //varTest += varStartDatePr + " " + varEndDatePr + "\r\n";

                        //throw new InvalidPluginExecutionException("This plugin is working\r\n" + varTest);
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

        // DEX date for year only
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

