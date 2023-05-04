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
	public class goRSAP: IPlugin
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
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters
                Entity entity = (Entity)context.InputParameters["Target"];

                // Verify that the target entity represents the Palm Go RSAP entity
                // If not, this plug-in was not registered correctly.
                // if (entity.LogicalName != "new_palmgorsap")
                //    return;

                try
                {
                    // Fixes:
                    // Remove ampersand from comments
                    // Compare values to Lower

                    // Global variables
                    string varDescription = ""; // Description field on form
                    bool varCreateExtract = false; // Create Extract field on form
                    string varProgram = ""; // Program field on form
                    Guid varRsapID = new Guid(); // GUID for palm go rsap record
                    StringBuilder sbHeaderList = new StringBuilder(); // Header for outreach strategies
                    StringBuilder sbHeaderList2 = new StringBuilder(); // Header for initial engagement
                    StringBuilder sbHeaderList3 = new StringBuilder(); // Header for support once housed
                    StringBuilder sbReportList = new StringBuilder(); // Data for outreach strategies
                    StringBuilder sbReportList2 = new StringBuilder(); // Data for initial engagement
                    StringBuilder sbReportList3 = new StringBuilder(); // Data for support once housed

                    // Create entity collection objects for fetchXML
                    EntityCollection result;
                    EntityCollection result2;
                    EntityCollection result3;
                    EntityCollection result4;
                    EntityCollection result5;
                    EntityCollection result6;

                    string varFileName = ""; // File name for outreach strategies
                    string varFileName2 = ""; // File name for initial engagement
                    string varFileName3 = ""; // File name for support once housed
                    DateTime varStartDate = new DateTime(); // Start date
                    DateTime varEndDate = new DateTime(); // End date
                    int varCheckInt = 0; // Parse integers
                    double varCheckDouble = 0; // Parse doubles
                    DateTime varCheckDate = new DateTime(); // Parse dates
                    EntityReference getEntity; // Entity reference object
                    AliasedValue getAlias; // Aliased value object

                    string varTest = ""; // Debug

                    // Only do this if the entity is the Palm Go RSAP entity
                    if (entity.LogicalName == "new_palmgorsap")
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

                        // Get info for current RSAP record
                        varDescription = entity.GetAttributeValue<string>("new_description");
                        varCreateExtract = entity.GetAttributeValue<bool>("new_createextract");
                        varStartDate = entity.GetAttributeValue<DateTime>("new_datefrom");
                        varEndDate = entity.GetAttributeValue<DateTime>("new_dateto");

                        // Get GUID
                        varRsapID = entity.Id;

                        // Get associated values for program
                        EntityReference ownerLookup = (EntityReference)entity.Attributes["new_program"];
                        varProgram += ownerLookup.Id.ToString() + ".\r\n";
                        varProgram += ((EntityReference)entity.Attributes["new_program"]).Name + ".\r\n";
                        varProgram += ownerLookup.LogicalName + ".\r\n";

                        var actualOwningUnit = _service.Retrieve(ownerLookup.LogicalName, ownerLookup.Id, new ColumnSet(true));

                        varProgram = actualOwningUnit["new_program"].ToString();

                        // Create file names for reports
                        varFileName = "AO Strategies " + varProgram + " " + varStartDate.ToString("MMM") + " " + varStartDate.Year + "-" + varEndDate.ToString("MMM") + " " + varEndDate.Year + ".xls";
                        varFileName2 = "Initial Engagement " + varProgram + " " + varStartDate.ToString("MMM") + " " + varStartDate.Year + "-" + varEndDate.ToString("MMM") + " " + varEndDate.Year + ".xls";
                        varFileName3 = "Support Once Housed" + varProgram + " " + varStartDate.ToString("MMM") + " " + varStartDate.Year + "-" + varEndDate.ToString("MMM") + " " + varEndDate.Year + ".xls";


                        // Part 1: AO Strategies

                        // Fetch statements for database
                        // Get the required fields from the AO Contact table (and associated entities)
                        // Any contacts aginst the program chosen
                        string dbClientList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmaocontact'>
                                    <attribute name='new_palmaocontactid' />
                                    <attribute name='new_name' />
                                    <attribute name='new_engagement' />
                                    <attribute name='new_entrydate' />
                                    <attribute name='new_disengaged' />
                                    <attribute name='new_lostserv' />
                                    <link-entity name='new_palmclient' to='new_client' from='new_palmclientid' link-type='outer'>
                                        <attribute name='new_palmclientid' />
                                        <attribute name='new_shorslk' />
                                    </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmaocontact' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmaocontact' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmaocontact' attribute='new_program' operator='eq' value='" + ownerLookup.Id + @"' />
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get total contacts for the period for each AO contact
                        string dbContactList = @"
                                <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                    <entity name='new_palmaooutreach'>
                                    <attribute name='new_aocontact' alias='total_count' aggregate='countcolumn'/> 
                                    <attribute name='new_aocontact' alias='total_count2' groupby='true'/> 
                                    <filter type='and'>
                                        <condition entityname='new_palmaooutreach' attribute='new_contdate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                        <condition entityname='new_palmaooutreach' attribute='new_contdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    </filter>
                                    </entity>
                                </fetch> ";

                        // Get total successful contacts for the period for each AO contact
                        string dbSuccessList = @"
                                <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                    <entity name='new_palmaooutreach'>
                                    <attribute name='new_aocontact' alias='total_count' aggregate='countcolumn'/> 
                                    <attribute name='new_aocontact' alias='total_count2' groupby='true'/> 
                                    <filter type='and'>
                                        <condition entityname='new_palmaooutreach' attribute='new_contdate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                        <condition entityname='new_palmaooutreach' attribute='new_contdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                        <condition entityname='new_palmaooutreach' attribute='new_successful' operator='eq' value='True' />
                                    </filter>
                                    </entity>
                                </fetch> ";

                        // Variables used to hold data returned from fetchXML (including sections below)
                        string dbPalmAOContactId = "";
                        string dbAlphaCode = "";
                        string dbId = "";
                        string dbEngagement = "";
                        string dbEntryDate = "";
                        string dbDisengaged = "";
                        string dbLostServ = "";
                        string dbAOContact = "";
                        string dbTotalCount = "";
                        string dbTotalCount2 = "";
                        string dbPalmClientId = "";
                        string dbDob = "";
                        string dbVulnScore = "";
                        string dbStartDate = "";
                        string dbEndDate = "";
                        string dbUnstable12 = "";
                        string dbUnstable25 = "";
                        string dbPriorHouse = "";
                        string dbIsFlexible = "No";
                        string dbWorkerAssess = "";
                        string dbClient = "";
                        string dbPlanDate = "";
                        string dbPlanType = "";
                        string dbPlanSummaryPhy = "";
                        string dbPlanSummaryMen = "";
                        string dbPlanSummarySus = "";
                        string dbPlanReason = "";
                        string dbPlanManage = "";
                        string dbIncludeRSAP = "";
                        string dbServAccessed = "";
                        string dbReferTo = "";
                        string dbIsHousing = "";
                        string dbReferDate = "";
                        string dbOwnerId = "";
                        string dbEmail = "";
                        string dbAccomType = "";
                        string dbTransType = "";
                        string dbStableType = "";
                        string dbSportAccess = "";
                        string dbSportStartDate = "";
                        string dbSportType = "";
                        string dbSportEndDate = "";
                        string dbEmployAccess = "";
                        string dbEmployStartDate = "";
                        string dbEmployType = "";
                        string dbEmployEndDate = "";
                        string dbEducationAccess = "";
                        string dbEducationStartDate = "";
                        string dbEducationType = "";
                        string dbEducationEndDate = "";
                        string dbActHealth = "";
                        string dbSupportServ = "";
                        string dbSupportFriend = "";
                        string dbDateFrom = "";
                        string dbDateTo = "";
                        string dbPlanStatus = "";
                        string dbDateServAccess = "";
                        string dbPeriod = "";
                        string dbSustain = "";
                        string dbFeelSystem = "";
                        string dbFeelFriend = "";
                        string dbAchieve = "";
                        string dbChallenge = "";
                        string dbIndependent = "";
                        string dbRiskTenancy = "";
                        string dbRiskType = "";
                        string dbRiskAction = "";

                        // Variables used for calculations or to translate the above values
                        bool varDoClient = false;
                        string varPrevClient = "";
                        string varPhysical = "No";
                        string varPhysicalDate = "";
                        string varPhysicalSumm = "";
                        string varPhysicalReason = "";
                        string varPhysicalManage = "";
                        string varMental = "No";
                        string varMentalDate = "";
                        string varMentalSumm = "";
                        string varMentalReason = "";
                        string varMentalManage = "";
                        string varBarriers = "No";
                        string varBarriersDate = "";
                        string varBarriersSumm = "";
                        string varServNeed = "";
                        string varServProvide = "";
                        string varServArrange = "";
                        string varServAccessedY = "No";
                        string varServAccessedNY = "No";
                        string varServAccessedN = "No";
                        string varStableHousing = "No";
                        string varDropOut = "No";
                        string varEmergAccom = "No";
                        string varTransAccom = "No";
                        string varTransType = "";
                        string varStable12 = "No";
                        string varStable25 = "No";
                        string varStableType = "";
                        string varReferTo = "";
                        string varReferDate = "";
                        string varWorker = "";
                        string varEmail = "";
                        string varBarriersOngoing = "";

                        string varTransDate = "";
                        string varAccomEnd = "";
                        string varPhysicalDate1 = "";
                        string varPhysicalSumm1 = "";
                        string varPhysicalReason1 = "";
                        string varPhysicalManage1 = "";
                        string varPhysicalDate2 = "";
                        string varPhysicalSumm2 = "";
                        string varPhysicalReason2 = "";
                        string varPhysicalManage2 = "";
                        string varMentalDate1 = "";
                        string varMentalSumm1 = "";
                        string varMentalReason1 = "";
                        string varMentalManage1 = "";
                        string varMentalDate2 = "";
                        string varMentalSumm2 = "";
                        string varMentalReason2 = "";
                        string varMentalManage2 = "";

                        string varSustain6 = "";
                        string varFeelSystem6 = "";
                        string varFeelFriend6 = "";
                        string varAchieve6 = "";
                        string varChallenge6 = "";
                        string varIndependent6 = "";
                        string varRiskTenancy6 = "";
                        string varRiskType6 = "";
                        string varRiskAction6 = "";
                        string varSustain12 = "";
                        string varFeelSystem12 = "";
                        string varFeelFriend12 = "";
                        string varAchieve12 = "";
                        string varChallenge12 = "";
                        string varIndependent12 = "";
                        string varRiskTenancy12 = "";
                        string varRiskType12 = "";
                        string varRiskAction12 = "";
                        string varSustain18 = "";
                        string varFeelSystem18 = "";
                        string varFeelFriend18 = "";
                        string varAchieve18 = "";
                        string varChallenge18 = "";
                        string varIndependent18 = "";
                        string varRiskTenancy18 = "";
                        string varRiskType18 = "";
                        string varRiskAction18 = "";
                        string varSustain24 = "";
                        string varFeelSystem24 = "";
                        string varFeelFriend24 = "";
                        string varAchieve24 = "";
                        string varChallenge24 = "";
                        string varIndependent24 = "";
                        string varRiskTenancy24 = "";
                        string varRiskType24 = "";
                        string varRiskAction24 = "";

                        // Timespan object for dates
                        TimeSpan ts1 = new TimeSpan();

                        // Get the fetch XML data and place in entity collection objects
                        result = _service.RetrieveMultiple(new FetchExpression(dbClientList));
                        result2 = _service.RetrieveMultiple(new FetchExpression(dbContactList));
                        result3 = _service.RetrieveMultiple(new FetchExpression(dbSuccessList));

                        // Loop through the AO contact data
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
                            if (c.FormattedValues.Contains("new_palmaocontactid"))
                                dbPalmAOContactId = c.FormattedValues["new_palmaocontactid"];
                            else if (c.Attributes.Contains("new_palmaocontactid"))
                                dbPalmAOContactId = c.Attributes["new_palmaocontactid"].ToString();
                            else
                                dbPalmAOContactId = "";

                            if (c.FormattedValues.Contains("new_name"))
                                dbId = c.FormattedValues["new_name"];
                            else if (c.Attributes.Contains("new_name"))
                                dbId = c.Attributes["new_name"].ToString();
                            else
                                dbId = "";

                            if (c.FormattedValues.Contains("new_palmclient1.new_shorslk"))
                                dbAlphaCode = c.FormattedValues["new_palmclient1.new_shorslk"];
                            else if (c.Attributes.Contains("new_palmclient1.new_shorslk"))
                                dbAlphaCode = c.GetAttributeValue<AliasedValue>("new_palmclient1.new_shorslk").Value.ToString();
                            else
                                dbAlphaCode = "";

                            if (c.FormattedValues.Contains("new_engagement"))
                                dbEngagement = c.FormattedValues["new_engagement"];
                            else if (c.Attributes.Contains("new_engagement"))
                                dbEngagement = c.Attributes["new_engagement"].ToString();
                            else
                                dbEngagement = "";

                            if (c.FormattedValues.Contains("new_entrydate"))
                                dbEntryDate = c.FormattedValues["new_entrydate"];
                            else if (c.Attributes.Contains("new_entrydate"))
                                dbEntryDate = c.Attributes["new_entrydate"].ToString();
                            else
                                dbEntryDate = "";

                            // Convert date from American format to Australian format
                            dbEntryDate = cleanDateAM(dbEntryDate);

                            if (c.FormattedValues.Contains("new_disengaged"))
                                dbDisengaged = c.FormattedValues["new_disengaged"];
                            else if (c.Attributes.Contains("new_disengaged"))
                                dbDisengaged = c.Attributes["new_disengaged"].ToString();
                            else
                                dbDisengaged = "";

                            if (c.FormattedValues.Contains("new_lostserv"))
                                dbLostServ = c.FormattedValues["new_lostserv"];
                            else if (c.Attributes.Contains("new_lostserv"))
                                dbLostServ = c.Attributes["new_lostserv"].ToString();
                            else
                                dbLostServ = "";

                            // Reset totals
                            dbTotalCount = "";
                            dbTotalCount2 = "";

                            // Loop through the total contacts data
                            foreach (var a in result2.Entities)
                            {

                                if (a.FormattedValues.Contains("total_count2"))
                                    dbAOContact = a.FormattedValues["total_count2"];
                                else if (a.Attributes.Contains("total_count2"))
                                    dbAOContact = a.Attributes["total_count2"].ToString();
                                else
                                    dbAOContact = "";

                                // Need to see if same contact id
                                if (dbId == dbAOContact)
                                {
                                    // Get the total
                                    if (a.FormattedValues.Contains("total_count"))
                                        dbTotalCount = a.FormattedValues["total_count"];
                                    else if (a.Attributes.Contains("total_count"))
                                        dbTotalCount = a.Attributes["total_count"].ToString();
                                    else
                                        dbTotalCount = "";

                                    dbTotalCount = cleanString(dbTotalCount, "double");
                                } // Same contact

                            } // Contact Loop

                            // Loop through the total successful contacts data
                            foreach (var a in result3.Entities)
                            {

                                if (a.FormattedValues.Contains("total_count2"))
                                    dbAOContact = a.FormattedValues["total_count2"];
                                else if (a.Attributes.Contains("total_count2"))
                                    dbAOContact = a.Attributes["total_count2"].ToString();
                                else
                                    dbAOContact = "";

                                // Need to see if same contact id
                                if (dbId == dbAOContact)
                                {
                                    // Get the total
                                    if (a.FormattedValues.Contains("total_count"))
                                        dbTotalCount2 = a.FormattedValues["total_count"];
                                    else if (a.Attributes.Contains("total_count"))
                                        dbTotalCount2 = a.Attributes["total_count"].ToString();
                                    else
                                        dbTotalCount2 = "";

                                    dbTotalCount2 = cleanString(dbTotalCount2, "double");
                                } // Same contact

                            } // Success Loop

                            // Append data to report
                            sbReportList.AppendLine("<tr>\r\n<td>" + dbId + "</td>\r\n<td>" + dbEntryDate + "</td>\r\n<td>" + dbTotalCount + "</td>\r\n<td>" + dbTotalCount2 + "</td>\r\n<td>" + dbEngagement + "</td>\r\n<td>" + dbDisengaged + "</td>\r\n<td>" + dbLostServ + "</td>\r\n<td>" + dbAlphaCode + "</td>\r\n</tr>");

                        } // client loop


                        //Header part of the RSAP extract
                        sbHeaderList.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                        sbHeaderList.AppendLine("<head>");
                        sbHeaderList.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                        sbHeaderList.AppendLine("<!--[if gte mso 9]>");
                        sbHeaderList.AppendLine("<xml>");
                        sbHeaderList.AppendLine("<x:ExcelWorkbook>");
                        sbHeaderList.AppendLine("<x:ExcelWorksheets>");
                        sbHeaderList.AppendLine("<x:ExcelWorksheet>");

                        //this line names the worksheet
                        sbHeaderList.AppendLine("<x:Name>Assertive outreach strategies</x:Name>");

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
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Alias</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Date of first outreach attempt (if known)</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Number of OUTREACH ATTEMPTS per client/ month</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Number of OUTREACH CONTACTS (i.e. client was approached and/or supported) per client/ month</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Current level of engagement as a RESULT of assertive outreach</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the person sleeping rough DISENGAGED since you/your team started working with this person (if applicable)?</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the rough sleeper been considered LOST TO SERVICES since you/your team started working with this person (if applicable)?</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Client's alphacode from SHIP (if/ when available)</td>");
                        sbHeaderList.AppendLine("</tr>");
                        sbHeaderList.AppendLine(sbReportList.ToString());
                        sbHeaderList.AppendLine("</table>");

                        //varTest += sbHeaderList.ToString();

                        // Part 2: Initial Engagement

                        // Fetch statements for database
                        // Get the required fields from the client table (and associated entities)
                        // Any clients that have support period ticked as RSAP
                        dbClientList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclient'>
                                <attribute name='new_palmclientid' />
                                <attribute name='new_address' />
                                <attribute name='new_shorslk' />
                                <attribute name='new_dob' />
                                <link-entity name='new_palmclientsupport' to='new_palmclientid' from='new_client' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_vulnscore' />
                                    <attribute name='new_startdate' />
                                    <attribute name='new_enddate' />
                                    <attribute name='new_unstable12' />
                                    <attribute name='new_unstable25' />
                                    <attribute name='new_priorhouse' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientsupport' attribute='new_rsap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_program' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    </filter >
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get the required fields from the financial table (and associated entities)
                        // Any data for the period against a support period ticked as RSAP
                        string dbFinancialList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientfinancial'>
                                <attribute name='new_palmclientfinancialid' />
                                <attribute name='new_isflexible' />
                                <attribute name='new_workerassess' />
                                <attribute name='new_supportperiod' />
                                <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_client' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientsupport' attribute='new_rsap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_program' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    </filter >
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get the required fields from the case plan table (and associated entities)
                        // Any data for the period against a support period ticked as RSAP
                        string dbCasePlanList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientcaseplan'>
                                <attribute name='new_palmclientcaseplanid' />
                                <attribute name='new_supportperiod' />
                                <link-entity name='new_palmclientcaseplanao' to='new_palmclientcaseplanid' from='new_caseplan' link-type='inner'>
                                    <attribute name='new_palmclientcaseplanaoid' />
                                    <attribute name='new_plandate' />
                                    <attribute name='new_plantype' />
                                    <attribute name='new_plansummaryphy' />
                                    <attribute name='new_plansummarymen' />
                                    <attribute name='new_plansummarysus' />
                                    <attribute name='new_planreason' />
                                    <attribute name='new_planmanage' />
                                </link-entity>
                                <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_client' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientcaseplanao' attribute='new_plandate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_rsap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_program' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    </filter >
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get the required fields from the services table (and associated entities)
                        // Any data for the period against a support period ticked as RSAP
                        string dbServicesList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientservices'>
                                <attribute name='new_palmclientservicesid' />
                                <attribute name='new_servneed' />
                                <attribute name='new_servprovide' />
                                <attribute name='new_servarrange' />
                                <attribute name='new_supportperiod' />
                                <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_client' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientservices' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientservices' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_rsap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_program' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    </filter >
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get the required fields from the refer to table (and associated entities)
                        // Any data for the period against a support period ticked as RSAP
                        string dbReferToList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientreferrals'>
                                <attribute name='new_palmclientreferralsid' />
                                <attribute name='new_includersap' />
                                <attribute name='new_servaccessed' />
                                <attribute name='new_referto' />
                                <attribute name='new_ishousing' />
                                <attribute name='new_entrydate' />
                                <attribute name='ownerid' />
                                <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_client' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientreferrals' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientreferrals' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientreferrals' attribute='new_includersap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_rsap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_program' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    </filter >
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get the required fields from the accom table (and associated entities)
                        // Any data for the period against a support period ticked as RSAP
                        string dbAccomList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientaccom'>
                                <attribute name='new_palmclientaccomid' />
                                <attribute name='new_datefrom' />
                                <attribute name='new_dateto' />
                                <attribute name='new_accomtype' />
                                <attribute name='new_transtype' />
                                <attribute name='new_stabletype' />
                                <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_client' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientsupport' attribute='new_rsap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_program' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    </filter >
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get the fetch XML data and place in entity collection objects
                        result = _service.RetrieveMultiple(new FetchExpression(dbClientList));
                        result2 = _service.RetrieveMultiple(new FetchExpression(dbFinancialList));
                        result3 = _service.RetrieveMultiple(new FetchExpression(dbCasePlanList));
                        result4 = _service.RetrieveMultiple(new FetchExpression(dbServicesList));
                        result5 = _service.RetrieveMultiple(new FetchExpression(dbReferToList));
                        result6 = _service.RetrieveMultiple(new FetchExpression(dbAccomList));

                        // Loop through client data
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

                            if (c.FormattedValues.Contains("new_shorslk"))
                                dbAlphaCode = c.FormattedValues["new_shorslk"];
                            else if (c.Attributes.Contains("new_shorslk"))
                                dbAlphaCode = c.Attributes["new_shorslk"].ToString();
                            else
                                dbAlphaCode = "";

                            if (c.FormattedValues.Contains("new_dob"))
                                dbDob = c.FormattedValues["new_dob"];
                            else if (c.Attributes.Contains("new_dob"))
                                dbDob = c.Attributes["new_dob"].ToString();
                            else
                                dbDob = "";

                            // Convert date from American format to Australian format
                            dbDob = cleanDateAM(dbDob);

                            varDoClient = false;
                            if (String.IsNullOrEmpty(varPrevClient) == true)
                                varDoClient = true;
                            else if (varPrevClient != dbPalmClientId)
                                varDoClient = true;

                            // Only do if this client is not the same as the previous client
                            if (varDoClient == true)
                            {
                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_vulnscore"))
                                    dbVulnScore = c.FormattedValues["new_palmclientsupport1.new_vulnscore"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_vulnscore"))
                                    dbVulnScore = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_vulnscore").Value.ToString();
                                else
                                    dbVulnScore = "";

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

                                if (String.IsNullOrEmpty(dbEndDate) == false)
                                    dbEndDate = "Yes";
                                else
                                    dbEndDate = "No";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_unstable12"))
                                    dbUnstable12 = c.FormattedValues["new_palmclientsupport1.new_unstable12"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_unstable12"))
                                    dbUnstable12 = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_unstable12").Value.ToString();
                                else
                                    dbUnstable12 = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_unstable25"))
                                    dbUnstable25 = c.FormattedValues["new_palmclientsupport1.new_unstable25"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_unstable25"))
                                    dbUnstable25 = c.GetAttributeValue <AliasedValue>("new_palmclientsupport1.new_unstable25").Value.ToString();
                                else
                                    dbUnstable25 = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_priorhouse"))
                                    dbPriorHouse = c.FormattedValues["new_palmclientsupport1.new_priorhouse"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_priorhouse"))
                                    dbPriorHouse = c.Attributes["new_palmclientsupport1.new_priorhouse"].ToString();
                                else
                                    dbPriorHouse = "";

                                //Reset variables
                                dbIsFlexible = "No";
                                dbWorkerAssess = "";
                                varPhysical = "No";
                                varPhysicalDate = "";
                                varPhysicalSumm = "";
                                varPhysicalReason = "";
                                varPhysicalManage = "";
                                varMental = "No";
                                varMentalDate = "";
                                varMentalSumm = "";
                                varMentalReason = "";
                                varMentalManage = "";
                                varBarriers = "No";
                                varBarriersDate = "";
                                varBarriersSumm = "";
                                varServNeed = "";
                                varServProvide = "";
                                varServArrange = "";
                                varServAccessedY = "No";
                                varServAccessedN = "No";
                                varStableHousing = "No";
                                varDropOut = "No";
                                varEmergAccom = "No";
                                varTransAccom = "No";
                                varTransType = "";
                                varStable12 = "No";
                                varStable25 = "No";
                                varStableType = "";
                                varReferTo = "";
                                varReferDate = "";
                                varWorker = "";
                                varEmail = "";

                                // Loop through financial data
                                foreach (var f in result2.Entities)
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

                                    // We need to get the entity id for the support period field for comparisons
                                    if (f.Attributes.Contains("new_palmclientsupport1.new_client"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)f.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_client").Value;
                                        dbClient = getEntity.Id.ToString();
                                    }
                                    else if (f.FormattedValues.Contains("new_palmclientsupport1.new_client"))
                                        dbClient = f.FormattedValues["new_palmclientsupport1.new_client"];
                                    else
                                        dbClient = "";

                                    // Only process if the client ids are the same
                                    if (dbClient == dbPalmClientId)
                                    {
                                        dbIsFlexible = "Yes"; // Flag as flexible

                                        // Get the worker assessment
                                        if (f.FormattedValues.Contains("new_workerassess"))
                                            dbWorkerAssess = f.FormattedValues["new_workerassess"];
                                        else if (f.Attributes.Contains("new_workerassess"))
                                            dbWorkerAssess = f.Attributes["new_workerassess"].ToString();
                                        else
                                            dbWorkerAssess = "";

                                        break;
                                    }
                                }

                                // Loop through case plan data
                                foreach (var p in result3.Entities)
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

                                    // We need to get the entity id for the support period field for comparisons
                                    if (p.Attributes.Contains("new_palmclientsupport2.new_client"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)p.GetAttributeValue<AliasedValue>("new_palmclientsupport2.new_client").Value;
                                        dbClient = getEntity.Id.ToString();
                                    }
                                    else if (p.FormattedValues.Contains("new_palmclientsupport2.new_client"))
                                        dbClient = p.FormattedValues["new_palmclientsupport2.new_client"];
                                    else
                                        dbClient = "";

                                    // Only process if the client ids are the same
                                    if (dbClient == dbPalmClientId)
                                    {
                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_plandate"))
                                            dbPlanDate = p.FormattedValues["new_palmclientcaseplanao1.new_plandate"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_plandate"))
                                            dbPlanDate = p.Attributes["new_palmclientcaseplanao1.new_plandate"].ToString();
                                        else
                                            dbPlanDate = "";

                                        // Convert date from American format to Australian format
                                        dbPlanDate = cleanDateAM(dbPlanDate);

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_plantype"))
                                            dbPlanType = p.FormattedValues["new_palmclientcaseplanao1.new_plantype"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_plantype"))
                                            dbPlanType = p.Attributes["new_palmclientcaseplanao1.new_plantype"].ToString();
                                        else
                                            dbPlanType = "";

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_plansummaryphy"))
                                            dbPlanSummaryPhy = p.FormattedValues["new_palmclientcaseplanao1.new_plansummaryphy"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_plansummaryphy"))
                                            dbPlanSummaryPhy = p.Attributes["new_palmclientcaseplanao1.new_plansummaryphy"].ToString();
                                        else
                                            dbPlanSummaryPhy = "";

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_plansummarymen"))
                                            dbPlanSummaryMen = p.FormattedValues["new_palmclientcaseplanao1.new_plansummarymen"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_plansummarymen"))
                                            dbPlanSummaryMen = p.Attributes["new_palmclientcaseplanao1.new_plansummarymen"].ToString();
                                        else
                                            dbPlanSummaryMen = "";

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_plansummarysus"))
                                            dbPlanSummarySus = p.FormattedValues["new_palmclientcaseplanao1.new_plansummarysus"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_plansummarysus"))
                                            dbPlanSummarySus = p.Attributes["new_palmclientcaseplanao1.new_plansummarysus"].ToString();
                                        else
                                            dbPlanSummarySus = "";

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_planreason"))
                                            dbPlanReason = p.FormattedValues["new_palmclientcaseplanao1.new_planreason"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_planreason"))
                                            dbPlanReason = p.Attributes["new_palmclientcaseplanao1.new_planreason"].ToString();
                                        else
                                            dbPlanReason = "";

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_planmanage"))
                                            dbPlanManage = p.FormattedValues["new_palmclientcaseplanao1.new_planmanage"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_planmanage"))
                                            dbPlanManage = p.Attributes["new_palmclientcaseplanao1.new_planmanage"].ToString();
                                        else
                                            dbPlanManage = "";

                                        // Get the details for the physical, mental and sustained housing case plans if not already done
                                        if (dbPlanType == "Physical" && String.IsNullOrEmpty(varPhysicalSumm) == false && String.IsNullOrEmpty(dbPlanSummaryPhy) == false)
                                        {
                                            varPhysical = "Yes";
                                            varPhysicalDate = dbPlanDate;
                                            varPhysicalSumm = dbPlanSummaryPhy;
                                            varPhysicalReason = dbPlanReason;
                                            varPhysicalManage = dbPlanManage;
                                        }
                                        else if (dbPlanType == "Mental" && String.IsNullOrEmpty(varMentalSumm) == false && String.IsNullOrEmpty(dbPlanSummaryMen) == false)
                                        {
                                            varMental = "Yes";
                                            varMentalDate = dbPlanDate;
                                            varMentalSumm = dbPlanSummaryMen;
                                            varMentalReason = dbPlanReason;
                                            varMentalManage = dbPlanManage;
                                        }
                                        else if (dbPlanType == "Sustained Housing" && String.IsNullOrEmpty(varBarriersSumm) == false && String.IsNullOrEmpty(dbPlanSummarySus) == false)
                                        {
                                            varBarriers = "Yes";
                                            varBarriersDate = dbPlanDate;
                                            varBarriersSumm = dbPlanSummarySus;
                                        }
                                    }
                                }

                                // Services no longer required

                                // Loop through refer to data
                                foreach (var r in result5.Entities)
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

                                    // We need to get the entity id for the client field for comparisons
                                    if (r.Attributes.Contains("new_palmclientsupport1.new_client"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)r.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_client").Value;
                                        dbClient = getEntity.Id.ToString();
                                    }
                                    else if (r.FormattedValues.Contains("new_palmclientsupport1.new_client"))
                                        dbClient = r.FormattedValues["new_palmclientsupport1.new_client"];
                                    else
                                        dbClient = "";

                                    // Only process if client ids are the same
                                    if (dbClient == dbPalmClientId)
                                    {
                                        // Force include in RSAP
                                        dbIncludeRSAP = "Yes";

                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (r.FormattedValues.Contains("new_servaccessed"))
                                            dbServAccessed = r.FormattedValues["new_servaccessed"];
                                        else if (r.Attributes.Contains("new_servaccessed"))
                                            dbServAccessed = r.Attributes["new_servaccessed"].ToString();
                                        else
                                            dbServAccessed = "";

                                        if (r.FormattedValues.Contains("new_referto"))
                                            dbReferTo = r.FormattedValues["new_referto"];
                                        else if (r.Attributes.Contains("new_referto"))
                                            dbReferTo = r.Attributes["new_referto"].ToString();
                                        else
                                            dbReferTo = "";

                                        if (r.FormattedValues.Contains("new_ishousing"))
                                            dbIsHousing = r.FormattedValues["new_ishousing"];
                                        else if (r.Attributes.Contains("new_ishousing"))
                                            dbIsHousing = r.Attributes["new_ishousing"].ToString();
                                        else
                                            dbIsHousing = "";

                                        if (r.FormattedValues.Contains("new_referdate"))
                                            dbReferDate = r.FormattedValues["new_referdate"];
                                        else if (r.Attributes.Contains("new_referdate"))
                                            dbReferDate = r.Attributes["new_referdate"].ToString();
                                        else
                                            dbReferDate = "";

                                        // Convert date from American format to Australian format
                                        dbReferDate = cleanDateAM(dbReferDate);

                                        if (r.FormattedValues.Contains("ownerid"))
                                            dbOwnerId = r.FormattedValues["ownerid"];
                                        else if (r.Attributes.Contains("ownerid"))
                                            dbOwnerId = r.Attributes["ownerid"].ToString();
                                        else
                                            dbOwnerId = "";

                                        dbEmail = "xx"; // Not currently working

                                        // Get service accessed values
                                        if (dbServAccessed == "Yes")
                                        {
                                            varServAccessedY = "Yes";
                                        }
                                        else if (dbServAccessed == "No")
                                        {
                                            varServAccessedN = "Yes";
                                        }

                                        // Get refer values if housing
                                        if (dbIsHousing == "Yes")
                                        {
                                            varReferTo = dbReferTo;
                                            varReferDate = dbReferDate;
                                            varWorker = dbOwnerId;
                                            varEmail = dbEmail;
                                        }

                                    }
                                }

                                // Loop through accommodation data
                                foreach (var a in result6.Entities)
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

                                    // We need to get the entity id for the support period field for comparisons
                                    if (a.Attributes.Contains("new_palmclientsupport1.new_client"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)a.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_client").Value;
                                        dbClient = getEntity.Id.ToString();
                                    }
                                    else if (a.FormattedValues.Contains("new_palmclientsupport1.new_client"))
                                        dbClient = a.FormattedValues["new_palmclientsupport1.new_client"];
                                    else
                                        dbClient = "";

                                    // Only process if support periods are the same
                                    if (dbClient == dbPalmClientId)
                                    {
                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (a.FormattedValues.Contains("new_datefrom"))
                                            dbDateFrom = a.FormattedValues["new_datefrom"];
                                        else if (a.Attributes.Contains("new_datefrom"))
                                            dbDateFrom = a.Attributes["new_datefrom"].ToString();
                                        else
                                            dbDateFrom = "";

                                        // Convert date from American format to Australian format
                                        dbDateFrom = cleanDateAM(dbDateFrom);

                                        if (a.FormattedValues.Contains("new_accomtype"))
                                            dbAccomType = a.FormattedValues["new_accomtype"];
                                        else if (a.Attributes.Contains("new_accomtype"))
                                            dbAccomType = a.Attributes["new_accomtype"].ToString();
                                        else
                                            dbAccomType = "";

                                        if (a.FormattedValues.Contains("new_transtype"))
                                            dbTransType = a.FormattedValues["new_transtype"];
                                        else if (a.Attributes.Contains("new_transtype"))
                                            dbTransType = a.Attributes["new_transtype"].ToString();
                                        else
                                            dbTransType = "";

                                        if (a.FormattedValues.Contains("new_stabletype"))
                                            dbStableType = a.FormattedValues["new_stabletype"];
                                        else if (a.Attributes.Contains("new_stabletype"))
                                            dbStableType = a.Attributes["new_stabletype"].ToString();
                                        else
                                            dbStableType = "";

                                        // Determine if stable housing was given during the timeframes
                                        if (String.IsNullOrEmpty(dbStableType) == false)
                                        {
                                            varStableHousing = "Yes";
                                            varStableType = dbStableType;

                                            if (String.IsNullOrEmpty(dbStartDate) == false && String.IsNullOrEmpty(dbDateFrom) == false)
                                            {
                                                ts1 = Convert.ToDateTime(dbDateFrom) - Convert.ToDateTime(dbStartDate);

                                                if (ts1.TotalDays < 84)
                                                {
                                                    varStable12 = "Yes";
                                                    dbUnstable12 = "";
                                                    dbUnstable25 = "";
                                                }
                                                else if (ts1.TotalDays < 175)
                                                {
                                                    varStable25 = "Yes";
                                                    dbUnstable25 = "";
                                                }
                                            }

                                        }

                                        // Get transitional type
                                        if (String.IsNullOrEmpty(dbTransType) == false)
                                        {
                                            varTransType = dbTransType;
                                            varTransAccom = "Yes";
                                        }

                                        // Determine housing provided
                                        if (dbAccomType == "Short term or emergency accommodation")
                                            varEmergAccom = "Yes";

                                        if (dbAccomType == "Medium term/transitional accommodation")
                                            varTransAccom = "Yes";
                                    }
                                }

                                // Determine if exited before stable housing
                                if (dbEndDate == "Yes" && varStableHousing == "No")
                                    varDropOut = "Yes";

                                // Add line to report
                                sbReportList2.AppendLine("<tr>\r\n<td>" + dbAlphaCode + "</td>\r\n<td>" + dbDob + "</td>\r\n<td>" + dbVulnScore + "</td>\r\n<td>" + dbIsFlexible + "</td>\r\n<td>" + dbWorkerAssess + "</td>\r\n<td>" + varPhysicalReason + "</td>\r\n<td>" + varPhysicalSumm + "</td>\r\n<td>" + varPhysicalManage + "</td>\r\n<td>" + varMentalReason + "</td>\r\n<td>" + varMentalSumm + "</td>\r\n<td>" + varMentalManage + "</td>\r\n<td>" + dbStartDate + "</td>\r\n<td>" + dbEndDate + "</td>\r\n<td>" + varBarriers + "</td>\r\n<td>" + varBarriersDate + "</td>\r\n<td>" + varBarriersSumm + "</td>\r\n<td>" + varEmergAccom + "</td>\r\n<td>" + varDropOut + "</td>\r\n<td>" + varTransAccom + "</td>\r\n<td>" + varTransType + "</td>\r\n<td>" + varStable12 + "</td>\r\n<td>" + dbUnstable12 + "</td>\r\n<td>" + varStable25 + "</td>\r\n<td>" + dbUnstable25 + "</td>\r\n<td>" + varStableType + "</td>\r\n<td>" + dbPriorHouse + "</td>\r\n<td>" + varReferTo + "</td>\r\n<td>" + varReferDate + "</td>\r\n<td>" + varWorker + "</td>\r\n\r\n<td>" + varEmail + "</td></tr>");

                            } //Do Client

                        } // Client Loop


                        //Header part of the RSAP extract
                        sbHeaderList2.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                        sbHeaderList2.AppendLine("<head>");
                        sbHeaderList2.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                        sbHeaderList2.AppendLine("<!--[if gte mso 9]>");
                        sbHeaderList2.AppendLine("<xml>");
                        sbHeaderList2.AppendLine("<x:ExcelWorkbook>");
                        sbHeaderList2.AppendLine("<x:ExcelWorksheets>");
                        sbHeaderList2.AppendLine("<x:ExcelWorksheet>");

                        //this line names the worksheet
                        sbHeaderList2.AppendLine("<x:Name>HRSAP Initial Engagement</x:Name>");

                        sbHeaderList2.AppendLine("<x:WorksheetOptions>");

                        sbHeaderList2.AppendLine("<x:Panes>");
                        sbHeaderList2.AppendLine("</x:Panes>");
                        sbHeaderList2.AppendLine("</x:WorksheetOptions>");
                        sbHeaderList2.AppendLine("</x:ExcelWorksheet>");
                        sbHeaderList2.AppendLine("</x:ExcelWorksheets>");
                        sbHeaderList2.AppendLine("</x:ExcelWorkbook>");
                        sbHeaderList2.AppendLine("</xml>");
                        sbHeaderList2.AppendLine("<![endif]-->");
                        sbHeaderList2.AppendLine("</head>");

                        sbHeaderList2.AppendLine("<table width=\"100%\" border=0 cellpadding=5 class=\"myClass1\">");
                        sbHeaderList2.AppendLine("<tr>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Alpha Code (SLK)</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Client's date of birth</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Client's vulnerability score (used to assess prioritisation)</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client received a 'flexible package'?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Worker's assessment of the impact of the flexible package on the client's needs and engagement with homelessness services</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has a PHYSICAL health care plan been developed by a health professional, based on the client's needs?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">PHYSICAL health concerns being addressed as part of the client's case plan</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Is the client MANAGING THE PHYSICAL HEALTH ISSUES identified as part of the healthcare plan/ case plan?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has a MENTAL health care plan been developed by a health professional, based on the client's needs?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">MENTAL health concerns being addressed as part of the client’s case plan</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Is the client MANAGING THE MENTAL HEALTH ISSUES identified as part of the healthcare plan/ case plan?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Date of Mobile Initial Assessment and Planning (IAP)</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client DISENGAGED from homelessness services after IAP?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Client has a case plan that includes support to sustain housing</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Date the case plan has been created</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Please specify the client's MAIN PERSONAL BARRIERS TO SUSTAIN HOUSING which are being focused on in the case plan</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client been provided crisis or emergency accommodation?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\"Has the client DISENGAGED from homelessness services before being offered transitional/ stable and suitable housing?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client been provided  transitional housing?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Type of transitional housing provided (if applicable)</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client been provided a stable and suitable housing WITHIN 12 WEEKS OF THEIR IAP?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">If not, what was the main reason for that?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client been provided a stable and suitable housing BETWEEN 12-25 WEEKS OF THEIR IAP?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">If not, what was the main reason for that?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">If yes, type of stable accommodation provided (if applicable)</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Did the client already have a housing offer/ tenancy prior to engagement?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Service provider to whom the client has been referred for HOUSING SUPPORT SERVICES</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Date the client was referred for housing support services?</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Name of the Caseworker/ Manager coordinating or responsible for the client's follow-up housing support services</td>");
                        sbHeaderList2.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Email address of the person responsible - Caseworker/ Manager</td>");
                        sbHeaderList2.AppendLine("</tr>");
                        sbHeaderList2.AppendLine(sbReportList2.ToString());
                        sbHeaderList2.AppendLine("</table>");


                        // Part 3: Support Once Housed

                        // Fetch statements for database
                        // Get the required fields from the client table (and associated entities)
                        // Any clients that have support period ticked as RSAP
                        dbClientList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclient'>
                                <attribute name='new_palmclientid' />
                                <attribute name='new_address' />
                                <attribute name='new_shorslk' />
                                <attribute name='new_dob' />
                                <link-entity name='new_palmclientsupport' to='new_palmclientid' from='new_client' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_vulnscore' />
                                    <attribute name='new_sportaccess' />
                                    <attribute name='new_sportstartdate' />
                                    <attribute name='new_sporttype' />
                                    <attribute name='new_sportenddate' />
                                    <attribute name='new_employaccess' />
                                    <attribute name='new_employstartdate' />
                                    <attribute name='new_employtype' />
                                    <attribute name='new_employenddate' />
                                    <attribute name='new_educationaccess' />
                                    <attribute name='new_educationstartdate' />
                                    <attribute name='new_educationtype' />
                                    <attribute name='new_educationenddate' />
                                    <attribute name='new_acthealth' />
                                    <attribute name='new_supportserv' />
                                    <attribute name='new_supportfriend' />
                                    <attribute name='new_priorhouse' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientsupport' attribute='new_rsap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_program' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    </filter >
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get the required fields from the accom table (and associated entities)
                        // Any data for the period against a support period ticked as RSAP
                        dbAccomList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientaccom'>
                                <attribute name='new_palmclientaccomid' />
                                <attribute name='new_datefrom' />
                                <attribute name='new_dateto' />
                                <attribute name='new_accomtype' />
                                <attribute name='new_transtype' />
                                <attribute name='new_stabletype' />
                                <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientsupport' attribute='new_rsap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_program' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    </filter >
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get the required fields from the case plan table (and associated entities)
                        // Any data for the period against a support period ticked as RSAP
                        dbCasePlanList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientcaseplan'>
                                <attribute name='new_palmclientcaseplanid' />
                                <attribute name='new_supportperiod' />
                                <link-entity name='new_palmclientcaseplanao' to='new_palmclientcaseplanid' from='new_caseplan' link-type='inner'>
                                    <attribute name='new_palmclientcaseplanaoid' />
                                    <attribute name='new_plandate' />
                                    <attribute name='new_plantype' />
                                    <attribute name='new_plansummaryphy' />
                                    <attribute name='new_plansummarymen' />
                                    <attribute name='new_plansummarysus' />
                                    <attribute name='new_planreason' />
                                    <attribute name='new_planmanage' />
                                    <attribute name='new_planstatus' />
                                </link-entity>
                                <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientcaseplanao' attribute='new_plandate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_rsap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_program' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    </filter >
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get the required fields from the refer to table (and associated entities)
                        // Any data for the period against a support period ticked as RSAP
                        dbReferToList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientreferrals'>
                                <attribute name='new_palmclientreferralsid' />
                                <attribute name='new_includersap' />
                                <attribute name='new_servaccessed' />
                                <attribute name='new_referto' />
                                <attribute name='new_ishousing' />
                                <attribute name='new_entrydate' />
                                <attribute name='new_dateservaccess' />
                                <attribute name='new_referto' />
                                <attribute name='ownerid' />
                                <link-entity name='new_palmclientsupport' to='new_supportperiod' from='new_palmclientsupportid' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientreferrals' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientreferrals' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientreferrals' attribute='new_includersap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_rsap' operator='eq' value='True' />
                                    <condition entityname='new_palmclientsupport' attribute='new_program' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_startdate' operator='lt' value='" + cleanDate(varEndDate) + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='null' />
                                        <condition entityname='new_palmclientsupport' attribute='new_enddate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    </filter >
                                </filter>
                              </entity>
                            </fetch> ";

                        // Get the required fields from the AO Scores table (and associated entities)
                        string dbAOScoresList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmaoscore'>
                                <attribute name='new_palmaoscoreid' />
                                <attribute name='new_period' />
                                <attribute name='new_sustain' />
                                <attribute name='new_feelsystem' />
                                <attribute name='new_feelfriend' />
                                <attribute name='new_achieve' />
                                <attribute name='new_challenge' />
                                <attribute name='new_independent' />
                                <attribute name='new_risktenancy' />
                                <attribute name='new_risktype' />
                                <attribute name='new_riskaction' />
                                <link-entity name='new_palmclient' to='new_client' from='new_palmclientid' link-type='inner'>
                                    <attribute name='new_palmclientid' />
                                </link-entity>
                              </entity>
                            </fetch> ";

                        // Get the fetch XML data and place in entity collection objects
                        result = _service.RetrieveMultiple(new FetchExpression(dbClientList));
                        result2 = _service.RetrieveMultiple(new FetchExpression(dbAccomList));
                        result3 = _service.RetrieveMultiple(new FetchExpression(dbCasePlanList));
                        result4 = _service.RetrieveMultiple(new FetchExpression(dbReferToList));
                        result5 = _service.RetrieveMultiple(new FetchExpression(dbAOScoresList));

                        // Loop through client data
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

                            if (c.FormattedValues.Contains("new_shorslk"))
                                dbAlphaCode = c.FormattedValues["new_shorslk"];
                            else if (c.Attributes.Contains("new_shorslk"))
                                dbAlphaCode = c.Attributes["new_shorslk"].ToString();
                            else
                                dbAlphaCode = "";

                            varDoClient = false;
                            if (String.IsNullOrEmpty(varPrevClient) == true)
                                varDoClient = true;
                            else if (varPrevClient != dbPalmClientId)
                                varDoClient = true;

                            // Only do if this client is not the same as the previous client
                            if (varDoClient == true)
                            {
                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_vulnscore"))
                                    dbVulnScore = c.FormattedValues["new_palmclientsupport1.new_vulnscore"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_vulnscore"))
                                    dbVulnScore = c.Attributes["new_palmclientsupport1.new_vulnscore"].ToString();
                                else
                                    dbVulnScore = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_sportaccess"))
                                    dbSportAccess = c.FormattedValues["new_palmclientsupport1.new_sportaccess"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_sportaccess"))
                                    dbSportAccess = c.Attributes["new_palmclientsupport1.new_sportaccess"].ToString();
                                else
                                    dbSportAccess = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_sportstartdate"))
                                    dbSportStartDate = c.FormattedValues["new_palmclientsupport1.new_sportstartdate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_sportstartdate"))
                                    dbSportStartDate = c.Attributes["new_palmclientsupport1.new_sportstartdate"].ToString();
                                else
                                    dbSportStartDate = "";

                                // Convert date from American format to Australian format
                                dbSportStartDate = cleanDateAM(dbSportStartDate);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_sporttype"))
                                    dbSportType = c.FormattedValues["new_palmclientsupport1.new_sporttype"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_sporttype"))
                                    dbSportType = c.Attributes["new_palmclientsupport1.new_sporttype"].ToString();
                                else
                                    dbSportType = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_sportenddate"))
                                    dbSportEndDate = c.FormattedValues["new_palmclientsupport1.new_sportenddate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_sportenddate"))
                                    dbSportEndDate = c.Attributes["new_palmclientsupport1.new_sportenddate"].ToString();
                                else
                                    dbSportEndDate = "";

                                // Convert date from American format to Australian format
                                dbSportEndDate = cleanDateAM(dbSportEndDate);

                                if (String.IsNullOrEmpty(dbSportEndDate) == false)
                                    dbSportEndDate = "Yes";
                                else
                                    dbSportEndDate = "No";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_employaccess"))
                                    dbEmployAccess = c.FormattedValues["new_palmclientsupport1.new_employaccess"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_employaccess"))
                                    dbEmployAccess = c.Attributes["new_palmclientsupport1.new_employaccess"].ToString();
                                else
                                    dbEmployAccess = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_employstartdate"))
                                    dbEmployStartDate = c.FormattedValues["new_palmclientsupport1.new_employstartdate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_employstartdate"))
                                    dbEmployStartDate = c.Attributes["new_palmclientsupport1.new_employstartdate"].ToString();
                                else
                                    dbEmployStartDate = "";

                                // Convert date from American format to Australian format
                                dbEmployStartDate = cleanDateAM(dbEmployStartDate);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_employtype"))
                                    dbEmployType = c.FormattedValues["new_palmclientsupport1.new_employtype"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_employtype"))
                                    dbEmployType = c.Attributes["new_palmclientsupport1.new_employtype"].ToString();
                                else
                                    dbEmployType = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_employenddate"))
                                    dbEmployEndDate = c.FormattedValues["new_palmclientsupport1.new_employenddate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_employenddate"))
                                    dbEmployEndDate = c.Attributes["new_palmclientsupport1.new_employenddate"].ToString();
                                else
                                    dbEmployEndDate = "";

                                // Convert date from American format to Australian format
                                dbEmployEndDate = cleanDateAM(dbEmployEndDate);

                                if (String.IsNullOrEmpty(dbEmployEndDate) == false)
                                    dbEmployEndDate = "Yes";
                                else
                                    dbEmployEndDate = "No";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_educationaccess"))
                                    dbEducationAccess = c.FormattedValues["new_palmclientsupport1.new_educationaccess"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_educationaccess"))
                                    dbEducationAccess = c.Attributes["new_palmclientsupport1.new_educationaccess"].ToString();
                                else
                                    dbEducationAccess = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_educationstartdate"))
                                    dbEducationStartDate = c.FormattedValues["new_palmclientsupport1.new_educationstartdate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_educationstartdate"))
                                    dbEducationStartDate = c.Attributes["new_palmclientsupport1.new_educationstartdate"].ToString();
                                else
                                    dbEducationStartDate = "";

                                // Convert date from American format to Australian format
                                dbEducationStartDate = cleanDateAM(dbEducationStartDate);

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_educationtype"))
                                    dbEducationType = c.FormattedValues["new_palmclientsupport1.new_educationtype"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_educationtype"))
                                    dbEducationType = c.Attributes["new_palmclientsupport1.new_educationtype"].ToString();
                                else
                                    dbEducationType = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_educationenddate"))
                                    dbEducationEndDate = c.FormattedValues["new_palmclientsupport1.new_educationenddate"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_educationenddate"))
                                    dbEducationEndDate = c.Attributes["new_palmclientsupport1.new_educationenddate"].ToString();
                                else
                                    dbEducationEndDate = "";

                                // Convert date from American format to Australian format
                                dbEducationEndDate = cleanDateAM(dbEducationEndDate);

                                if (String.IsNullOrEmpty(dbEducationEndDate) == false)
                                    dbEducationEndDate = "Yes";
                                else
                                    dbEducationEndDate = "No";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_acthealth"))
                                    dbActHealth = c.FormattedValues["new_palmclientsupport1.new_acthealth"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_acthealth"))
                                    dbActHealth = c.Attributes["new_palmclientsupport1.new_acthealth"].ToString();
                                else
                                    dbActHealth = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supportserv"))
                                    dbSupportServ = c.FormattedValues["new_palmclientsupport1.new_supportserv"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supportserv"))
                                    dbSupportServ = c.Attributes["new_palmclientsupport1.new_supportserv"].ToString();
                                else
                                    dbSupportServ = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_supportfriend"))
                                    dbSupportFriend = c.FormattedValues["new_palmclientsupport1.new_supportfriend"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_supportfriend"))
                                    dbSupportFriend = c.Attributes["new_palmclientsupport1.new_supportfriend"].ToString();
                                else
                                    dbSupportFriend = "";

                                if (c.FormattedValues.Contains("new_palmclientsupport1.new_priorhouse"))
                                    dbPriorHouse = c.FormattedValues["new_palmclientsupport1.new_priorhouse"];
                                else if (c.Attributes.Contains("new_palmclientsupport1.new_priorhouse"))
                                    dbPriorHouse = c.Attributes["new_palmclientsupport1.new_priorhouse"].ToString();
                                else
                                    dbPriorHouse = "";

                                // Reset variables
                                varTransType = "";
                                varStableType = "";
                                varTransDate = "";
                                varAccomEnd = "";

                                varPhysicalDate = "";
                                varPhysicalSumm = "";
                                varPhysicalReason = "";
                                varPhysicalManage = "";
                                varPhysicalDate1 = "";
                                varPhysicalSumm1 = "";
                                varPhysicalReason1 = "";
                                varPhysicalManage1 = "";
                                varPhysicalDate2 = "";
                                varPhysicalSumm2 = "";
                                varPhysicalReason2 = "";
                                varPhysicalManage2 = "";
                                varMentalDate = "";
                                varMentalSumm = "";
                                varMentalReason = "";
                                varMentalManage = "";
                                varMentalDate1 = "";
                                varMentalSumm1 = "";
                                varMentalReason1 = "";
                                varMentalManage1 = "";
                                varMentalDate2 = "";
                                varMentalSumm2 = "";
                                varMentalReason2 = "";
                                varMentalManage2 = "";

                                varSustain6 = "";
                                varFeelSystem6 = "";
                                varFeelFriend6 = "";
                                varAchieve6 = "";
                                varChallenge6 = "";
                                varIndependent6 = "";
                                varRiskTenancy6 = "";
                                varRiskType6 = "";
                                varRiskAction6 = "";
                                varSustain12 = "";
                                varFeelSystem12 = "";
                                varFeelFriend12 = "";
                                varAchieve12 = "";
                                varChallenge12 = "";
                                varIndependent12 = "";
                                varRiskTenancy12 = "";
                                varRiskType12 = "";
                                varRiskAction12 = "";
                                varSustain18 = "";
                                varFeelSystem18 = "";
                                varFeelFriend18 = "";
                                varAchieve18 = "";
                                varChallenge18 = "";
                                varIndependent18 = "";
                                varRiskTenancy18 = "";
                                varRiskType18 = "";
                                varRiskAction18 = "";
                                varSustain24 = "";
                                varFeelSystem24 = "";
                                varFeelFriend24 = "";
                                varAchieve24 = "";
                                varChallenge24 = "";
                                varIndependent24 = "";
                                varRiskTenancy24 = "";
                                varRiskType24 = "";
                                varRiskAction24 = "";

                                // Loop through accom data
                                foreach (var a in result2.Entities)
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

                                    // We need to get the entity id for the client field for comparisons
                                    if (a.Attributes.Contains("new_palmclientsupport1.new_client"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)a.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_client").Value;
                                        dbClient = getEntity.Id.ToString();
                                    }
                                    else if (a.FormattedValues.Contains("new_palmclientsupport1.new_client"))
                                        dbClient = a.FormattedValues["new_palmclientsupport1.new_client"];
                                    else
                                        dbClient = "";

                                    // Only do if client ids are the same
                                    if (dbClient == dbPalmClientId)
                                    {
                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (a.FormattedValues.Contains("new_datefrom"))
                                            dbDateFrom = a.FormattedValues["new_datefrom"];
                                        else if (a.Attributes.Contains("new_datefrom"))
                                            dbDateFrom = a.Attributes["new_datefrom"].ToString();
                                        else
                                            dbDateFrom = "";

                                        // Convert date from American format to Australian format
                                        dbDateFrom = cleanDateAM(dbDateFrom);

                                        if (a.FormattedValues.Contains("new_accomtype"))
                                            dbAccomType = a.FormattedValues["new_accomtype"];
                                        else if (a.Attributes.Contains("new_accomtype"))
                                            dbAccomType = a.Attributes["new_accomtype"].ToString();
                                        else
                                            dbAccomType = "";

                                        if (a.FormattedValues.Contains("new_transtype"))
                                            dbTransType = a.FormattedValues["new_transtype"];
                                        else if (a.Attributes.Contains("new_transtype"))
                                            dbTransType = a.Attributes["new_transtype"].ToString();
                                        else
                                            dbTransType = "";

                                        if (a.FormattedValues.Contains("new_stabletype"))
                                            dbStableType = a.FormattedValues["new_stabletype"];
                                        else if (a.Attributes.Contains("new_stabletype"))
                                            dbStableType = a.Attributes["new_stabletype"].ToString();
                                        else
                                            dbStableType = "";

                                        if (a.FormattedValues.Contains("new_dateto"))
                                            dbDateTo = a.FormattedValues["new_dateto"];
                                        else if (a.Attributes.Contains("new_dateto"))
                                            dbDateTo = a.Attributes["new_dateto"].ToString();
                                        else
                                            dbDateTo = "";

                                        // Convert date from American format to Australian format
                                        dbDateTo = cleanDateAM(dbDateTo);

                                        // Get dates for stable accom
                                        if (String.IsNullOrEmpty(dbStableType) == false)
                                        {
                                            varStableType = dbStableType;

                                            if (String.IsNullOrEmpty(dbDateFrom) == false)
                                            {
                                                if (String.IsNullOrEmpty(varTransDate) == true)
                                                    varTransDate = dbDateFrom;
                                                else if (Convert.ToDateTime(varTransDate) > Convert.ToDateTime(dbDateFrom))
                                                    varTransDate = dbDateFrom;
                                            }

                                        }

                                        // Get transitional type
                                        if (String.IsNullOrEmpty(dbTransType) == false)
                                        {
                                            varTransType = dbTransType;
                                        }

                                        // Get accom end
                                        if ((String.IsNullOrEmpty(dbStableType) == false || String.IsNullOrEmpty(dbTransType) == false) && String.IsNullOrEmpty(dbDateTo) == false)
                                            varAccomEnd = dbDateTo;
                                    }
                                }

                                // Loop through case plan data
                                foreach (var p in result3.Entities)
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

                                    // We need to get the entity id for the client field for comparisons
                                    if (p.Attributes.Contains("new_palmclientsupport2.new_client"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)p.GetAttributeValue<AliasedValue>("new_palmclientsupport2.new_client").Value;
                                        dbClient = getEntity.Id.ToString();
                                    }
                                    else if (p.FormattedValues.Contains("new_palmclientsupport2.new_client"))
                                        dbClient = p.FormattedValues["new_palmclientsupport2.new_client"];
                                    else
                                        dbClient = "";

                                    // Only process if client ids are the same
                                    if (dbClient == dbPalmClientId)
                                    {
                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_plandate"))
                                            dbPlanDate = p.FormattedValues["new_palmclientcaseplanao1.new_plandate"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_plandate"))
                                            dbPlanDate = p.Attributes["new_palmclientcaseplanao1.new_plandate"].ToString();
                                        else
                                            dbPlanDate = "";

                                        // Convert date from American format to Australian format
                                        dbPlanDate = cleanDateAM(dbPlanDate);

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_plantype"))
                                            dbPlanType = p.FormattedValues["new_palmclientcaseplanao1.new_plantype"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_plantype"))
                                            dbPlanType = p.Attributes["new_palmclientcaseplanao1.new_plantype"].ToString();
                                        else
                                            dbPlanType = "";

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_plansummaryphy"))
                                            dbPlanSummaryPhy = p.FormattedValues["new_palmclientcaseplanao1.new_plansummaryphy"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_plansummaryphy"))
                                            dbPlanSummaryPhy = p.Attributes["new_palmclientcaseplanao1.new_plansummaryphy"].ToString();
                                        else
                                            dbPlanSummaryPhy = "";

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_plansummarymen"))
                                            dbPlanSummaryMen = p.FormattedValues["new_palmclientcaseplanao1.new_plansummarymen"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_plansummarymen"))
                                            dbPlanSummaryMen = p.Attributes["new_palmclientcaseplanao1.new_plansummarymen"].ToString();
                                        else
                                            dbPlanSummaryMen = "";

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_plansummarysus"))
                                            dbPlanSummarySus = p.FormattedValues["new_palmclientcaseplanao1.new_plansummarysus"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_plansummarysus"))
                                            dbPlanSummarySus = p.Attributes["new_palmclientcaseplanao1.new_plansummarysus"].ToString();
                                        else
                                            dbPlanSummarySus = "";

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_planreason"))
                                            dbPlanReason = p.FormattedValues["new_palmclientcaseplanao1.new_planreason"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_planreason"))
                                            dbPlanReason = p.Attributes["new_palmclientcaseplanao1.new_planreason"].ToString();
                                        else
                                            dbPlanReason = "";

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_planmanage"))
                                            dbPlanManage = p.FormattedValues["new_palmclientcaseplanao1.new_planmanage"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_planmanage"))
                                            dbPlanManage = p.Attributes["new_palmclientcaseplanao1.new_planmanage"].ToString();
                                        else
                                            dbPlanManage = "";

                                        if (p.FormattedValues.Contains("new_palmclientcaseplanao1.new_planstatus"))
                                            dbPlanStatus = p.FormattedValues["new_palmclientcaseplanao1.new_planstatus"];
                                        else if (p.Attributes.Contains("new_palmclientcaseplanao1.new_planstatus"))
                                            dbPlanStatus = p.Attributes["new_palmclientcaseplanao1.new_planstatus"].ToString();
                                        else
                                            dbPlanStatus = "";

                                        // Get the physical, mental and sustained housing case plan data for current and previous plans
                                        if (dbPlanType == "Physical" && String.IsNullOrEmpty(dbPlanSummaryPhy) == false)
                                        {

                                            if (String.IsNullOrEmpty(varPhysicalSumm) == true)
                                            {
                                                varPhysicalDate = dbPlanDate;
                                                varPhysicalSumm = dbPlanSummaryPhy;
                                                varPhysicalReason = dbPlanReason;
                                                varPhysicalManage = dbPlanManage;
                                            }
                                            else if (String.IsNullOrEmpty(varPhysicalSumm1) == true)
                                            {
                                                varPhysicalDate1 = dbPlanDate;
                                                varPhysicalSumm1 = dbPlanSummaryPhy;
                                                varPhysicalReason1 = dbPlanReason;
                                                varPhysicalManage1 = dbPlanManage;
                                            }
                                            else if (String.IsNullOrEmpty(varPhysicalSumm2) == true)
                                            {
                                                varPhysicalDate2 = dbPlanDate;
                                                varPhysicalSumm2 = dbPlanSummaryPhy;
                                                varPhysicalReason2 = dbPlanReason;
                                                varPhysicalManage2 = dbPlanManage;
                                            }

                                        }
                                        else if (dbPlanType == "Mental" && String.IsNullOrEmpty(dbPlanSummaryMen) == false)
                                        {

                                            if (String.IsNullOrEmpty(varMentalSumm) == true)
                                            {
                                                varMentalDate = dbPlanDate;
                                                varMentalSumm = dbPlanSummaryMen;
                                                varMentalReason = dbPlanReason;
                                                varMentalManage = dbPlanManage;
                                            }
                                            else if (String.IsNullOrEmpty(varMentalSumm1) == true)
                                            {
                                                varMentalDate1 = dbPlanDate;
                                                varMentalSumm1 = dbPlanSummaryMen;
                                                varMentalReason1 = dbPlanReason;
                                                varMentalManage1 = dbPlanManage;
                                            }
                                            else if (String.IsNullOrEmpty(varMentalSumm2) == true)
                                            {
                                                varMentalDate2 = dbPlanDate;
                                                varMentalSumm2 = dbPlanSummaryMen;
                                                varMentalReason2 = dbPlanReason;
                                                varMentalManage2 = dbPlanManage;
                                            }

                                        }
                                        else if (dbPlanType == "Sustained Housing" && String.IsNullOrEmpty(dbPlanSummarySus) == false)
                                        {

                                            if (String.IsNullOrEmpty(varBarriersSumm) == true)
                                            {
                                                varBarriers = "Yes";
                                                varBarriersSumm = dbPlanSummarySus;

                                                if (dbPlanStatus == "Ongoing")
                                                    varBarriersOngoing = "Yes";
                                                else
                                                    varBarriersOngoing = "No";
                                            }

                                        }
                                    }
                                }

                                // Loop through refer to rows
                                foreach (var r in result4.Entities)
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

                                    // We need to get the entity id for the client field for comparisons
                                    if (r.Attributes.Contains("new_palmclientsupport1.new_client"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)r.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_client").Value;
                                        dbClient = getEntity.Id.ToString();
                                    }
                                    else if (r.FormattedValues.Contains("new_palmclientsupport1.new_client"))
                                        dbClient = r.FormattedValues["new_palmclientsupport1.new_client"];
                                    else
                                        dbClient = "";

                                    // Only process if client ids are the same
                                    if (dbClient == dbPalmClientId)
                                    {
                                        dbIncludeRSAP = "Yes"; // Default

                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (r.FormattedValues.Contains("new_servaccessed"))
                                            dbServAccessed = r.FormattedValues["new_servaccessed"];
                                        else if (r.Attributes.Contains("new_servaccessed"))
                                            dbServAccessed = r.Attributes["new_servaccessed"].ToString();
                                        else
                                            dbServAccessed = "";

                                        if (r.FormattedValues.Contains("new_referto"))
                                            dbReferTo = r.FormattedValues["new_referto"];
                                        else if (r.Attributes.Contains("new_referto"))
                                            dbReferTo = r.Attributes["new_referto"].ToString();
                                        else
                                            dbReferTo = "";

                                        if (r.FormattedValues.Contains("new_ishousing"))
                                            dbIsHousing = r.FormattedValues["new_ishousing"];
                                        else if (r.Attributes.Contains("new_ishousing"))
                                            dbIsHousing = r.Attributes["new_ishousing"].ToString();
                                        else
                                            dbIsHousing = "";

                                        if (r.FormattedValues.Contains("new_referdate"))
                                            dbReferDate = r.FormattedValues["new_referdate"];
                                        else if (r.Attributes.Contains("new_referdate"))
                                            dbReferDate = r.Attributes["new_referdate"].ToString();
                                        else
                                            dbReferDate = "";

                                        // Convert date from American format to Australian format
                                        dbReferDate = cleanDateAM(dbReferDate);

                                        if (r.FormattedValues.Contains("new_dateservaccess"))
                                            dbDateServAccess = r.FormattedValues["new_dateservaccess"];
                                        else if (r.Attributes.Contains("new_dateservaccess"))
                                            dbDateServAccess = r.Attributes["new_dateservaccess"].ToString();
                                        else
                                            dbDateServAccess = "";

                                        // Convert date from American format to Australian format
                                        dbDateServAccess = cleanDateAM(dbDateServAccess);

                                        if (r.FormattedValues.Contains("ownerid"))
                                            dbOwnerId = r.FormattedValues["ownerid"];
                                        else if (r.Attributes.Contains("ownerid"))
                                            dbOwnerId = r.Attributes["ownerid"].ToString();
                                        else
                                            dbOwnerId = "";

                                        dbEmail = "xx"; // Need to get from user table

                                        // Service accessed information
                                        if (dbServAccessed == "Yes")
                                        {
                                            varServAccessedY = "Yes";

                                            if (String.IsNullOrEmpty(dbEntryDate) == false && String.IsNullOrEmpty(dbDateServAccess) == false)
                                            {
                                                if (dbEntryDate != dbDateServAccess)
                                                    varServAccessedNY = "Yes";
                                            }

                                        }

                                    }
                                }

                                // Loop through support period data
                                foreach (var s in result5.Entities)
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

                                    // We need to get the entity id for the client field for comparisons
                                    if (s.Attributes.Contains("new_palmclientsupport2.new_client"))
                                    {
                                        // Get the entity id for the client using the entity reference object
                                        getEntity = (EntityReference)s.GetAttributeValue<AliasedValue>("new_palmclientsupport2.new_client").Value;
                                        dbClient = getEntity.Id.ToString();
                                    }
                                    else if (s.FormattedValues.Contains("new_palmclientsupport2.new_client"))
                                        dbClient = s.FormattedValues["new_palmclientsupport2.new_client"];
                                    else
                                        dbClient = "";

                                    // Only do if client ids are the same
                                    if (dbClient == dbPalmClientId)
                                    {
                                        // Process the data as follows:
                                        // If there is a formatted value for the field, use it
                                        // Otherwise if there is a literal value for the field, use it
                                        // Otherwise the value wasn't returned so set as nothing
                                        if (s.FormattedValues.Contains("new_period"))
                                            dbPeriod = s.FormattedValues["new_period"];
                                        else if (s.Attributes.Contains("new_period"))
                                            dbPeriod = s.Attributes["new_period"].ToString();
                                        else
                                            dbPeriod = "";

                                        if (s.FormattedValues.Contains("new_sustain"))
                                            dbSustain = s.FormattedValues["new_sustain"];
                                        else if (s.Attributes.Contains("new_sustain"))
                                            dbSustain = s.Attributes["new_sustain"].ToString();
                                        else
                                            dbSustain = "";

                                        if (s.FormattedValues.Contains("new_feelsystem"))
                                            dbFeelSystem = s.FormattedValues["new_feelsystem"];
                                        else if (s.Attributes.Contains("new_feelsystem"))
                                            dbFeelSystem = s.Attributes["new_feelsystem"].ToString();
                                        else
                                            dbFeelSystem = "";

                                        if (s.FormattedValues.Contains("new_feelfriend"))
                                            dbFeelFriend = s.FormattedValues["new_feelfriend"];
                                        else if (s.Attributes.Contains("new_feelfriend"))
                                            dbFeelFriend = s.Attributes["new_feelfriend"].ToString();
                                        else
                                            dbFeelFriend = "";

                                        if (s.FormattedValues.Contains("new_achieve"))
                                            dbAchieve = s.FormattedValues["new_achieve"];
                                        else if (s.Attributes.Contains("new_achieve"))
                                            dbAchieve = s.Attributes["new_achieve"].ToString();
                                        else
                                            dbAchieve = "";

                                        if (s.FormattedValues.Contains("new_challenge"))
                                            dbChallenge = s.FormattedValues["new_challenge"];
                                        else if (s.Attributes.Contains("new_challenge"))
                                            dbChallenge = s.Attributes["new_challenge"].ToString();
                                        else
                                            dbChallenge = "";

                                        if (s.FormattedValues.Contains("new_independent"))
                                            dbIndependent = s.FormattedValues["new_independent"];
                                        else if (s.Attributes.Contains("new_independent"))
                                            dbIndependent = s.Attributes["new_independent"].ToString();
                                        else
                                            dbIndependent = "";

                                        if (s.FormattedValues.Contains("new_risktenancy"))
                                            dbRiskTenancy = s.FormattedValues["new_risktenancy"];
                                        else if (s.Attributes.Contains("new_risktenancy"))
                                            dbRiskTenancy = s.Attributes["new_risktenancy"].ToString();
                                        else
                                            dbRiskTenancy = "";

                                        if (s.FormattedValues.Contains("new_risktype"))
                                            dbRiskType = s.FormattedValues["new_risktype"];
                                        else if (s.Attributes.Contains("new_risktype"))
                                            dbRiskType = s.Attributes["new_risktype"].ToString();
                                        else
                                            dbRiskType = "";

                                        if (s.FormattedValues.Contains("new_riskaction"))
                                            dbRiskAction = s.FormattedValues["new_riskaction"];
                                        else if (s.Attributes.Contains("new_riskaction"))
                                            dbRiskAction = s.Attributes["new_riskaction"].ToString();
                                        else
                                            dbRiskAction = "";

                                        // Get the information based on the period
                                        if (dbPeriod == "6")
                                        {
                                            varSustain6 = dbSustain;
                                            varFeelSystem6 = dbFeelSystem;
                                            varFeelFriend6 = dbFeelFriend;
                                            varAchieve6 = dbAchieve;
                                            varChallenge6 = dbChallenge;
                                            varIndependent6 = dbIndependent;
                                            varRiskTenancy6 = dbRiskTenancy;
                                            varRiskType6 = dbRiskType;
                                            varRiskAction6 = dbRiskAction;
                                        }
                                        else if (dbPeriod == "12")
                                        {
                                            varSustain12 = dbSustain;
                                            varFeelSystem12 = dbFeelSystem;
                                            varFeelFriend12 = dbFeelFriend;
                                            varAchieve12 = dbAchieve;
                                            varChallenge12 = dbChallenge;
                                            varIndependent12 = dbIndependent;
                                            varRiskTenancy12 = dbRiskTenancy;
                                            varRiskType12 = dbRiskType;
                                            varRiskAction12 = dbRiskAction;
                                        }
                                        else if (dbPeriod == "18")
                                        {
                                            varSustain18 = dbSustain;
                                            varFeelSystem18 = dbFeelSystem;
                                            varFeelFriend18 = dbFeelFriend;
                                            varAchieve18 = dbAchieve;
                                            varChallenge18 = dbChallenge;
                                            varIndependent18 = dbIndependent;
                                            varRiskTenancy18 = dbRiskTenancy;
                                            varRiskType18 = dbRiskType;
                                            varRiskAction18 = dbRiskAction;
                                        }
                                        else if (dbPeriod == "24")
                                        {
                                            varSustain24 = dbSustain;
                                            varFeelSystem24 = dbFeelSystem;
                                            varFeelFriend24 = dbFeelFriend;
                                            varAchieve24 = dbAchieve;
                                            varChallenge24 = dbChallenge;
                                            varIndependent24 = dbIndependent;
                                            varRiskTenancy24 = dbRiskTenancy;
                                            varRiskType24 = dbRiskType;
                                            varRiskAction24 = dbRiskAction;
                                        }
                                    }
                                }

                                // Add to report row
                                sbReportList3.AppendLine("<tr>\r\n<td>" + dbAlphaCode + "</td>\r\n<td>" + dbVulnScore + "</td>\r\n<td>" + varTransType + "</td>\r\n<td>" + varStableType + "</td>\r\n<td>" + dbPriorHouse + "</td>\r\n<td>" + varTransDate + "</td>\r\n<td>" + varPhysicalReason1 + "</td>\r\n<td>" + varPhysicalSumm1 + "</td>\r\n<td>" + varPhysicalManage1 + "</td>\r\n<td>" + varMentalReason1 + "</td>\r\n<td>" + varMentalSumm1 + "</td>\r\n<td>" + varMentalManage1 + "</td>\r\n<td>" + varBarriers + "</td>\r\n<td>" + varBarriersSumm + "</td>\r\n<td>" + dbSupportServ + "</td>\r\n<td>" + dbSupportFriend + "</td>\r\n<td>" + dbSportType + "</td>\r\n<td>" + dbSportEndDate + "</td>\r\n<td>" + dbEmployType + "</td>\r\n<td>" + dbEmployEndDate + "</td>\r\n<td>" + dbActHealth + "</td>\r\n<td>" + varSustain6 + "</td>\r\n<td>" + varRiskTenancy6 + "</td>\r\n<td>" + varRiskType6 + "</td>\r\n<td>" + varRiskAction6 + "</td>\r\n<td>" + varFeelSystem6 + "</td>\r\n<td>" + varFeelFriend6 + "</td>\r\n<td>" + varAchieve6 + "</td>\r\n<td>" + varChallenge6 + "</td>\r\n<td>" + varIndependent6 + "</td>\r\n<td>" + varSustain12 + "</td>\r\n<td>" + varRiskTenancy12 + "</td>\r\n<td>" + varRiskType12 + "</td>\r\n<td>" + varRiskAction12 + "</td>\r\n<td>" + varFeelSystem12 + "</td>\r\n<td>" + varFeelFriend12 + "</td>\r\n<td>" + varAchieve12 + "</td>\r\n<td>" + varChallenge12 + "</td>\r\n<td>" + varIndependent12 + "</td>\r\n<td>" + varSustain18 + "</td>\r\n<td>" + varRiskTenancy18 + "</td>\r\n<td>" + varRiskType18 + "</td>\r\n<td>" + varRiskAction18 + "</td>\r\n<td>" + varFeelSystem18 + "</td>\r\n<td>" + varFeelFriend18 + "</td>\r\n<td>" + varAchieve18 + "</td>\r\n<td>" + varChallenge18 + "</td>\r\n<td>" + varIndependent18 + "</td>\r\n<td>" + varSustain24 + "</td>\r\n<td>" + varRiskTenancy24 + "</td>\r\n<td>" + varRiskType24 + "</td>\r\n<td>" + varRiskAction24 + "</td>\r\n<td>" + varFeelSystem24 + "</td>\r\n<td>" + varFeelFriend24 + "</td>\r\n<td>" + varAchieve24 + "</td>\r\n<td>" + varChallenge24 + "</td>\r\n<td>" + varIndependent24 + "</td>\r\n<td>" + varAccomEnd + "</td>\r\n</tr>");

                            } //doclient
                        }


                        //Header part of the RSAP extract
                        sbHeaderList3.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                        sbHeaderList3.AppendLine("<head>");
                        sbHeaderList3.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                        sbHeaderList3.AppendLine("<!--[if gte mso 9]>");
                        sbHeaderList3.AppendLine("<xml>");
                        sbHeaderList3.AppendLine("<x:ExcelWorkbook>");
                        sbHeaderList3.AppendLine("<x:ExcelWorksheets>");
                        sbHeaderList3.AppendLine("<x:ExcelWorksheet>");

                        //this line names the worksheet
                        sbHeaderList3.AppendLine("<x:Name>HRSAP Support Once Housed</x:Name>");

                        sbHeaderList3.AppendLine("<x:WorksheetOptions>");

                        sbHeaderList3.AppendLine("<x:Panes>");
                        sbHeaderList3.AppendLine("</x:Panes>");
                        sbHeaderList3.AppendLine("</x:WorksheetOptions>");
                        sbHeaderList3.AppendLine("</x:ExcelWorksheet>");
                        sbHeaderList3.AppendLine("</x:ExcelWorksheets>");
                        sbHeaderList3.AppendLine("</x:ExcelWorkbook>");
                        sbHeaderList3.AppendLine("</xml>");
                        sbHeaderList3.AppendLine("<![endif]-->");
                        sbHeaderList3.AppendLine("</head>");

                        sbHeaderList3.AppendLine("<table width=\"100%\" border=0 cellpadding=5 class=\"myClass1\">");
                        sbHeaderList3.AppendLine("<tr>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Alpha Code(SLK)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Client's vulnerability score (used to assess prioritisation)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Type of transitional housing provided to the client (if applicable)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Type of stable accommodation provided (if applicable)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Did the client already have a housing offer/ tenancy prior to engagement?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Date in which the client entered transitional or stable housing (TENANCY START DATE)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has a PHYSICAL health care plan been developed by a health professional, based on the client's needs?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">PHYSICAL health concerns being addressed as part of the client's case plan</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Is the client MANAGING THE PHYSICAL HEALTH ISSUES identified as part of the healthcare plan/ case plan?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has a MENTAL health care plan been developed by a health professional, based on the client's needs?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">MENTAL health concerns being addressed as part of the client’s case plan</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Is the client MANAGING THE MENTAL HEALTH ISSUES identified as part of the healthcare plan/ case plan?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Client has a case plan that includes addressing any barriers to housing</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Please specify the client's MAIN PERSONAL BARRIERS TO SUSTAIN HOUSING which are being focused on in the case plan</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How supported the client feels in managing their housing status BY THE SERVICE SYSTEM [BASELINE]</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How supported the client feels in managing their housing status BY FAMILY, FRIENDS, COMMUNITY [BASELINE]</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">What is the main type of community-based/ sport/ recreational activity being accessed by the client?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client STOPPED accessing community-based/ sport/ recreational activities (if known)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Please describe the type of employment activities</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client STOPPED accessing employment and/or education-related activities (if known)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How are these activities impacting the client's social inclusion and economic participation?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client sustained transitional or stable housing up to 6 months?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How would you currently assess the tenancy risk since the client moved to transitional/ stable housing 6 months ago?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Type of issue which led to the tenancy being at risk in this reporting period (if applicable)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Actions pursued to prevent tenancy risk or eviction</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How supported the client feels in managing their housing status BY THE SERVICE SYSTEM [REASSESSMENT] </td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How supported the client feels in managing their housing status BY FAMILY, FRIENDS, COMMUNITY [REASSESSMENT]</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Main ACHIEVEMENTS of the client by sustaining long-term housing (6 months)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Main CHALLENGES faced by the client to sustain long-term housing (6 months)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client MOVED TO AN INDEPENDENT HOUSING status?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client sustained transitional or stable housing up to 12 months?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How would you currently assess the tenancy risk since the client moved to transitional/ stable housing 12 months ago?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Type of issue which led to the tenancy being at risk in this reporting period (if applicable)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Actions pursued to prevent tenancy risk or eviction</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How supported the client feels in managing their housing status BY THE SERVICE SYSTEM [REASSESSMENT]</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How supported the client feels in managing their housing status BY FAMILY, FRIENDS, COMMUNITY [REASSESSMENT]</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Main ACHIEVEMENTS of the client by sustaining long-term housing (12 months)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Main CHALLENGES faced by the client to sustain long-term housing (12 months)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client MOVED TO AN INDEPENDENT HOUSING status?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client sustained transitional or stable housing up to 18 months?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How would you currently assess the tenancy risk since the client moved to transitional/ stable housing 18 months ago?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Type of issue which led to the tenancy being at risk in this reporting period (if applicable)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Actions pursued to prevent tenancy risk or eviction</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How supported the client feels in managing their housing status BY THE SERVICE SYSTEM [REASSESSMENT]</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How supported the client feels in managing their housing status BY FAMILY, FRIENDS, COMMUNITY [REASSESSMENT]</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Main ACHIEVEMENTS of the client by sustaining long-term housing (18 months)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Main CHALLENGES faced by the client to sustain long-term housing (18 months)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client MOVED TO AN INDEPENDENT HOUSING status?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client sustained transitional or stable housing up to 24 months?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How would you currently assess the tenancy risk since the client moved to transitional/ stable housing 24 months ago?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Type of issue which led to the tenancy being at risk in this reporting period (if applicable)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Actions pursued to prevent tenancy risk or eviction</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How supported the client feels in managing their housing status BY THE SERVICE SYSTEM [REASSESSMENT]</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">How supported the client feels in managing their housing status BY FAMILY, FRIENDS, COMMUNITY [REASSESSMENT]</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Main ACHIEVEMENTS of the client by sustaining long-term housing (24 months)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Main CHALLENGES faced by the client to sustain long-term housing (24 months)</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Has the client MOVED TO AN INDEPENDENT HOUSING status?</td>");
                        sbHeaderList3.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Date in which the client left transitional or stable housing (TENANCY END DATE, if known)</td>");
                        sbHeaderList3.AppendLine("</tr>");
                        sbHeaderList3.AppendLine(sbReportList3.ToString());
                        sbHeaderList3.AppendLine("</table>");

                        // Create note against current Palm Go RSAP record and add attachment
                        // AO Strategies
                        byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                        string encodedData = System.Convert.ToBase64String(filename);
                        Entity Annotation = new Entity("annotation");
                        Annotation.Attributes["objectid"] = new EntityReference("new_palmgorsap", varRsapID);
                        Annotation.Attributes["objecttypecode"] = "new_palmgorsap";
                        Annotation.Attributes["subject"] = "RSAP Extract";
                        Annotation.Attributes["documentbody"] = encodedData;
                        Annotation.Attributes["mimetype"] = @"application/msexcel";
                        Annotation.Attributes["notetext"] = "AO Strategies " + cleanDate(varStartDate) + " to " + cleanDate(varEndDate);
                        Annotation.Attributes["filename"] = varFileName;
                        _service.Create(Annotation);

                        // Initial Engagement
                        byte[] filename2 = Encoding.ASCII.GetBytes(sbHeaderList2.ToString());
                        string encodedData2 = System.Convert.ToBase64String(filename2);
                        Entity Annotation2 = new Entity("annotation");
                        Annotation2.Attributes["objectid"] = new EntityReference("new_palmgorsap", varRsapID);
                        Annotation2.Attributes["objecttypecode"] = "new_palmgorsap";
                        Annotation2.Attributes["subject"] = "RSAP Extract";
                        Annotation2.Attributes["documentbody"] = encodedData2;
                        Annotation2.Attributes["mimetype"] = @"application/msexcel";
                        Annotation2.Attributes["notetext"] = "Initial Engagement " + cleanDate(varStartDate) + " to " + cleanDate(varEndDate);
                        Annotation2.Attributes["filename"] = varFileName2;
                        _service.Create(Annotation2);

                        // Support once housed
                        byte[] filename3 = Encoding.ASCII.GetBytes(sbHeaderList3.ToString());
                        string encodedData3 = System.Convert.ToBase64String(filename3);
                        Entity Annotation3 = new Entity("annotation");
                        Annotation2.Attributes["objectid"] = new EntityReference("new_palmgorsap", varRsapID);
                        Annotation2.Attributes["objecttypecode"] = "new_palmgorsap";
                        Annotation2.Attributes["subject"] = "RSAP Extract";
                        Annotation2.Attributes["documentbody"] = encodedData3;
                        Annotation2.Attributes["mimetype"] = @"application/msexcel";
                        Annotation2.Attributes["notetext"] = "Support Once Housed " + cleanDate(varStartDate) + " to " + cleanDate(varEndDate);
                        Annotation2.Attributes["filename"] = varFileName3;
                        _service.Create(Annotation3);

                        // throw new InvalidPluginExecutionException("This plugin is working\r\n" + varTest);
                    }


                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in the FollowupPlugin plug-in.", ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("FollowupPlugin - Issue here: {0}", ex.ToString());
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

        // DEX format date for year
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

