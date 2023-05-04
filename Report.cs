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
    public class goReport : IPlugin
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
                //</snippetFollowupPlugin2>

                // Verify that the target entity represents the Palm Go Report entity
                // If not, this plug-in was not registered correctly.
                // if (entity.LogicalName != "new_palmgoreport")
                //    return;

                try
                {
                    // Fixes:
                    // Remove ampersand from comments
                    // Compare values to Lower


                    string varDescription = ""; // Report Description
                    string varReport = ""; // Report Type
                    Guid varReportId = new Guid(); // GUID for report 
                    StringBuilder sbHeaderList = new StringBuilder(); // String builder for header
                    StringBuilder sbReportList = new StringBuilder(); // String builder for report
                    StringBuilder sbFinancialList = new StringBuilder(); // String builder for financial list
                    StringBuilder sbErrorList = new StringBuilder(); // String builder for errors

                    string varFileName = "No Report"; // Report file name
                    string varReportName = "No Report";

                    DateTime varStartDate = new DateTime(); // Report start date
                    DateTime varEndDate = new DateTime(); // Report end date
                    DateTime varStartDatePr = new DateTime(); // Print start date
                    DateTime varEndDatePr = new DateTime(); // Print end date



                    // Variables for data
                    string dbFinList = "";
                    string dbSupportList = "";
                    string dbSupportList2 = "";
                    string dbSupportList3 = "";
                    string dbActivitiesList = "";
                    string dbReferralsList = "";
              


                    //NEW
                    Int32 aggregateSupport = 0;
                    Int32 newSupport = 0;
                    Int32 closedSupport = 0;
                    Double activitiesSum = 0;
                    Int32 refCount = 0;
                    Int32 refCountIn = 0;
                    Int32 refCountEx = 0;
                    string refInternalCount = "";
                    string refExternalCount = "";
                    Int32 totalNewSupport = 0;
                    Int32 totalClosedSupport = 0;
                    Int32 totalActiveSupport = 0;

                    Int32 aggregateFin = 0;
                    Int32 totalAggregateFin = 0;
                    Double sumFin = 0;
                    Double totalSumFin = 0;

                    Int32 countEmergency = 0;
                    Int32 countRentArrears = 0;
                    Int32 countRentAdvance = 0;
                    Int32 countBonds = 0;
                    Int32 countMaterial = 0;
                    Double sumEmergency = 0;
                    Double sumRentArrears = 0;
                    Double sumRentAdvance = 0;
                    Double sumBonds = 0;
                    Double sumMaterial = 0;

                    string clientTotals = "";
                    string KPIMalleeString = "";
                    string KPIMalleeActivities = "";
                    string KPIMalleeFinancials = "";

                    var iapCentral = new Record[2];
                    iapCentral[0] = new Record("{A85305DE-088B-E811-A9C5-000D3AD1C8A7}", "IAP - Bdgo", "new_palmsuprogram");
                    iapCentral[1] = new Record("{AC5305DE-088B-E811-A9C5-000D3AD1C8A7}", "IAP - Pris", "new_palmsuprogram");
                    var iapMallee = new Record[1];
                    iapMallee[0] = new Record("{AA5305DE-088B-E811-A9C5-000D3AD1C8A7}", "IAP - Mlda", "new_palmsuprogram");
                    var iapMetro = new Record[1];
                    iapMetro[0] = new Record("{FC551085-FD50-E911-A976-000D3AD1C904}", "IAP - Preston", "new_palmsuprogram");
                    Record[][] iapPrograms = {iapCentral, iapMallee, iapMetro};


                    var sfarCentral = new Record[1];
                    sfarCentral[0] = new Record("{E45305DE-088B-E811-A9C5-000D3AD1C8A7}", "SFAR - Bdgo", "new_palmsuprogram");
                    var sfarMallee = new Record[1];
                    sfarMallee[0] = new Record("{E65305DE-088B-E811-A9C5-000D3AD1C8A7}", "SFAR - Mlda", "new_palmsuprogram");
                    var sfarMetro = new Record[0];
                    Record[][] sfarPrograms = { sfarCentral, sfarMallee, sfarMetro };

                    var prapCentral = new Record[2];
                    prapCentral[0] = new Record("{5E5305DE-088B-E811-A9C5-000D3AD1C8A7}", "PRAP - Bdgo","new_palmsuprogram");
                    prapCentral[1] = new Record("{37FB3B22-0726-EB11-A813-000D3AD1CF6F}", "PRAP Plus Bgo", "new_palmsuprogram");
                    var prapMallee = new Record[2];
                    prapMallee[0] = new Record("{2CFB3B22-0726-EB11-A813-000D3AD1CF6F}", "PRAP - Mallee", "new_palmsuprogram");
                    prapMallee[1] = new Record("{979F4E28-0726-EB11-A813-000D3AD1CF6F}", "PRAP Plus Mallee","new_palmsuprogram");
                    var prapMetro = new Record[0];
                    Record[][] prapPrograms = { prapCentral, prapMallee, prapMetro };

                    var saapBendigo = new Record[1];
                    saapBendigo[0] = new Record("{8C5305DE-088B-E811-A9C5-000D3AD1C8A7}", "HCM - Bdgo", "new_palmsuprogram");
                    Record[][] saapBdgoProgram = { saapBendigo };

                    var saapKyabram = new Record[1];
                    saapKyabram[0] = new Record("{8E5305DE-088B-E811-A9C5-000D3AD1C8A7}", "HCM - Kybm", "new_palmsuprogram");
                    Record[][] saapKyabPrograms = { saapKyabram };

                    var icmiCentral = new Record[1];
                    icmiCentral[0] = new Record("{B25305DE-088B-E811-A9C5-000D3AD1C8A7}", "ICMI - Echa", "new_palmsuprogram");
                    var icmiMallee = new Record[0];
                    var icmiMetro = new Record[0];
                    Record[][] icmiPrograms = { icmiCentral, icmiMallee, icmiMetro };

                    var rsapCentral = new Record[1];
                    rsapCentral[0] = new Record("{7AEF5F2E-0726-EB11-A813-000D3AD1CF6F}", "RSAP AO Bgo", "new_palmsuprogram");
                    var rsapMallee = new Record[1];
                    rsapMallee[0] = new Record("{769ABC34-0726-EB11-A813-000D3AD1CF6F}", "RSAP AO SHill", "new_palmsuprogram");
                    var rsapMetro = new Record[0];
                    Record[][] rsapPrograms = { rsapCentral, rsapMallee, rsapMetro };

                    var hsdCentral = new Record[1];
                    hsdCentral[0] = new Record("{7C9ABC34-0726-EB11-A813-000D3AD1CF6F}", "Housing Direct", "new_palmsuprogram"); //??
                    var hsdMallee = new Record[0];
                    var hsdMetro = new Record[0];
                    Record[][] hsdPrograms = { hsdCentral, hsdMallee, hsdMetro };

                    var taapCentral = new Record[0];
                    var taapMallee = new Record[1];
                    taapMallee[0] = new Record("","","");
                    var taapMetro = new Record[0];
                    Record[][] taapPrograms = { taapCentral, taapMallee, taapMetro };

                    var tenPlusCentral = new Record[0];
                    var tenPlusMallee = new Record[1];
                    tenPlusMallee[0] = new Record("{625305DE-088B-E811-A9C5-000D3AD1C8A7}", "Tenancy Plus", "new_palmsuprogram");
                    var tenPlusMetro = new Record[0];
                    Record[][] tenPlusPrograms = { tenPlusCentral, tenPlusMallee, tenPlusMetro};

                    //Financials Area
                    var hefCentral = new Record[4];
                    hefCentral[0] = new Record("{B8BFA9E2-9B9A-E811-A9CA-000D3AD1CE4E}", "HEF - Lodn", "new_palmsubrokerage");
                    hefCentral[1] = new Record("{BABFA9E2-9B9A-E811-A9CA-000D3AD1CE4E}", "HEF - HCM", "new_palmsubrokerage");
                    hefCentral[2] = new Record("{BCBFA9E2-9B9A-E811-A9CA-000D3AD1CE4E}", "HEF - Kybm", "new_palmsubrokerage");
                    hefCentral[3] = new Record("{BEBFA9E2-9B9A-E811-A9CA-000D3AD1CE4E}", "HEF - Pris", "new_palmsubrokerage");
                    var hefMallee = new Record[1];
                    hefMallee[0] = new Record("{C4BFA9E2-9B9A-E811-A9CA-000D3AD1CE4E}", "HEF - Mall", "new_palmsubrokerage");
                    var hefMetro = new Record[1];
                    hefMetro[0] = new Record("{59EDB982-5608-EB11-A813-000D3ACAADC0}", "HEF - Preston", "new_palmsubrokerage");
                    Record[][] hefFinancials = { hefCentral, hefMallee, hefMetro };

                    var erCentral = new Record[0];
                    var erMallee = new Record[2];
                    erMallee[0] = new Record("{36C0A9E2-9B9A-E811-A9CA-000D3AD1CE4E}", "ER - N/W", "new_palmsubrokerage");
                    erMallee[1] = new Record("{34C0A9E2-9B9A-E811-A9CA-000D3AD1CE4E}", "ER - Mury", "new_palmsubrokerage");
                    var erMetro = new Record[0];
                    Record[][] erFinanicals = { erCentral, erMallee, erMetro };

                    var prapFinCentral = new Record[0];
                    var prapFinMallee = new Record[1];
                    prapFinMallee[0] = new Record("{A4BFA9E2-9B9A-E811-A9CA-000D3AD1CE4E}","PRAP - Mallee", "new_palmsubrokerage");
                    var prapFinMetro = new Record[0];
                    Record[][] prapFinancials = { prapFinCentral, prapFinMallee, prapFinMetro };


                    Record hefFin = new Record("", "", "new_palmsubrokerage");

                    // Entity collection objects for fetch XML data
                    EntityCollection result;
                    EntityCollection result2;
                    EntityCollection result3;
                    EntityCollection result4;
                    EntityCollection result5;
                    EntityCollection result6;


                    // Only do this if the entity is the Palm Go reports entity
                    if (entity.LogicalName == "new_connectgoreport")
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

                        // Get info for current Financial record
                        varDescription = entity.GetAttributeValue<string>("new_description");
                        varStartDate = entity.GetAttributeValue<DateTime>("new_datefrom");
                        varStartDatePr = Convert.ToDateTime(varStartDate.AddHours(14).ToString()); // Australian Date

                        varEndDate = entity.GetAttributeValue<DateTime>("new_dateto");
                        varEndDatePr = Convert.ToDateTime(varEndDate.AddHours(23).ToString()); // Australian Date

                        varEndDate = varEndDate.AddHours(23); // Correct for Australian time

                        // Get the report type
                        if (entity.Contains("new_reporttype"))
                            varReport = entity.FormattedValues["new_reporttype"];



                        Record[][][] programsSDT = { saapBdgoProgram, saapKyabPrograms, icmiPrograms, rsapPrograms, hsdPrograms, iapPrograms, sfarPrograms, prapPrograms };
                        string[] programNamesSDT = new [] { "Bendigo SAAP", "Kyabram SAAP", "ICMI", "RSAP", "Housing Direct", "IAP","SFAR", "PRAP" };


                        Record[] programsKPIMallee = { iapMallee[0], prapMallee[0], prapMallee[1], tenPlusMallee[0]};
                        string[] programNamesKPIMallee = new [] { "IAP Mildura", "PRAP Mallee", "PRAP Plus Mildura", "Tenancy Plus Mildura" };

                        Record[] financialKPIMallee = { hefMallee[0], erMallee[0], erMallee[1], prapFinMallee[0] }; //Contains information on brokerages
                        string[] financialNamesKPIMallee = new[] {"HEF - Mallee","ER - North West", "ER - Murray", "PRAP/Plus - Mallee" }; //Contains names of the brokerages
                        string[] financialFilterKPIMallee = new[] {
                            "<condition attribute='new_financialtype' operator='eq' value='100000000' /><condition attribute='new_assistance' operator='eq' uiname='Accommodation' uitype='new_palmsuassistance' value='{37F595ED-9B9A-E811-A9C0-000D3AD1C695}' /><condition attribute='new_shor' operator='eq' value='100000009' />", 
                            "<condition attribute='new_financialtype' operator='eq' value='100000000' /><condition attribute='new_assistance' operator='eq' uiname='Rent Arrears' uitype='new_palmsuassistance' value='{5BF595ED-9B9A-E811-A9C0-000D3AD1C695}' />",
                            "<condition attribute='new_financialtype' operator='eq' value='100000000' /><condition attribute='new_assistance' operator='eq' uiname='Rent In Advance' uitype='new_palmsuassistance' value='{5DF595ED-9B9A-E811-A9C0-000D3AD1C695}' />",
                            "<condition attribute='new_financialtype' operator='eq' value='100000000' /><condition attribute='new_assistance' operator='eq' uiname='Bonds' uitype='new_palmsuassistance' value='{3DF595ED-9B9A-E811-A9C0-000D3AD1C695}' />",
                            "<condition attribute='new_financialtype' operator='eq' value='100000000' /><filter type='or'><condition attribute='new_assistance' operator='in'>" +
                            "<value uiname='Aids/Equipment' uitype='new_palmsuassistance'>{39F595ED-9B9A-E811-A9C0-000D3AD1C695}</value>" +
                            "<value uiname='Food' uitype='new_palmsuassistance'>{43F595ED-9B9A-E811-A9C0-000D3AD1C695}</value>" +
                            "<value uiname='Miscellaneous' uitype='new_palmsuassistance'>{4FF595ED-9B9A-E811-A9C0-000D3AD1C695}</value>" +
                            "<value uiname='Travel' uitype='new_palmsuassistance'>{63F595ED-9B9A-E811-A9C0-000D3AD1C695}</value></condition></filter>"
                        }; //Contains values used to filter assistance type - {Short term, Rent in arr, Rent in advance, Bonds, Material Aid}


                        // Get the entity ID
                        varReportId = entity.Id;

                        // This data is processed for thebudget report
                        if (varReport == "Service Delivery Tracking Information")
                        {

                            // Get the report names and file names
                            varReportName = "Service Delivery Tracking Information - " + varStartDatePr.ToString("MMM") + " " + varStartDatePr.Year + " to " + varEndDatePr.ToString("MMM") + " " + varEndDatePr.Year;
                            varFileName = "Service Delivery Tracking Information - " + varStartDatePr.ToString("MMM") + " " + varStartDatePr.Year + " to " + varEndDatePr.ToString("MMM") + " " + varEndDatePr.Year + ".xls";

                            clientTotals = "<tr><td>New Support Periods</td><td><b>Central</b></td><td><b>Mallee</b></td><td><b>Metro</b></td></tr>";

                            for (var x = 0; x < programsSDT.Length; x++)
                            {
                                clientTotals += "<tr><td><b>" + programNamesSDT[x] + "</b></td>";

                                for (var i = 0; i < programsSDT[x].Length; i++)
                                {
                                    aggregateSupport = 0;
                                    if (programsSDT[x][i].Length != 0)
                                    {

                                        for (var j = 0; j < programsSDT[x][i].Length; j++)
                                        {
                                            //Get the support periods created within the selected period of time.
                                            dbSupportList = @"
                                            <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                                <entity name='new_palmclientsupport'>
                                                <attribute name='new_palmclientsupportid' alias='support_count' aggregate='count' />
                                                <filter type='and'>
                                                    <condition attribute='new_startdate' operator='on-or-after' value='" + varStartDate + @"' />
                                                    <condition attribute='new_startdate' operator='on-or-before' value='" + varEndDate + @"' />
                                                    <condition attribute='new_program' operator='eq' uiname='" + programsSDT[x][i][j].recName + "' uitype='" + programsSDT[x][i][j].theType + "' value='" + programsSDT[x][i][j].guid + @"' />
                                                </filter>
                                                </entity>
                                            </fetch>";

                                            // Get the fetch XML data and place in entity collection objects
                                            result = _service.RetrieveMultiple(new FetchExpression(dbSupportList));

                                            // Support Period Numbers
                                            foreach (var c in result.Entities)
                                            {
                                                aggregateSupport += (Int32)((AliasedValue)c["support_count"]).Value;
                                            }
                                        }
                                        clientTotals += "<td>" + aggregateSupport + "</td>";
                                    }
                                    else
                                    {
                                        clientTotals += "<td>N/A</td>";
                                    }
                                }
                            }
                            clientTotals += "</tr>";

                            clientTotals += "<tr><td>&nbsp;</td></tr>";

                            clientTotals += "<tr><td>Financial Records</td><td><b>Central</b></td><td><b>Mallee</b></td><td><b>Metro</b></td></tr><tr><td><b>HEF Fin</b></td>";
                            for (int i = 0; i < hefFinancials.Length; i++)
                            {
                                aggregateFin = 0;
                                for (int j = 0; j < hefFinancials[i].Length; j++)
                                {
                                    //Get the financials created within the selected period of time.
                                    dbFinList = @"
                                    <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                        <entity name='new_palmclientfinancial'>
                                            <attribute name='new_palmclientfinancialid' alias='fin_count' aggregate='count' />
                                            <filter type='and'>
                                                <condition attribute='new_entrydate' operator='on-or-after' value='" + varStartDate + @"' />
                                                <condition attribute='new_entrydate' operator='on-or-before' value='" + varEndDate + @"' />
                                                <condition attribute='new_brokerage' operator='eq' uiname='" + hefFinancials[i][j].recName + "' uitype='" + hefFinancials[i][j].theType + "' value='" + hefFinancials[i][j].guid + @"' />
                                            </filter>
                                        </entity>
                                    </fetch>";

                                    // Get the fetch XML data and place in entity collection objects
                                    result2 = _service.RetrieveMultiple(new FetchExpression(dbFinList));


                                    // Client Details
                                    foreach (var c in result2.Entities)
                                    {
                                        aggregateFin += (Int32)((AliasedValue)c["fin_count"]).Value;
                                    }

                                }
                                clientTotals += "<td>" + aggregateFin + "</td>"; //Add each Central, Mallee, Metro to row.
                            }
                            clientTotals += "</tr>";

                            //Header part of the summary report
                            sbHeaderList.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                            sbHeaderList.AppendLine("<head>");
                            sbHeaderList.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                            sbHeaderList.AppendLine("<!--[if gte mso 9]>");
                            sbHeaderList.AppendLine("<xml>");
                            sbHeaderList.AppendLine("<x:ExcelWorkbook>");
                            sbHeaderList.AppendLine("<x:ExcelWorksheets>");
                            sbHeaderList.AppendLine("<x:ExcelWorksheet>");

                            //this line names the worksheet
                            sbHeaderList.AppendLine("<x:Name>Report Data</x:Name>");

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

                            // Produce the summary using the values above
                            sbHeaderList.AppendLine("<table border=0 cellpadding=3>");
                            sbHeaderList.AppendLine("<tr><td colspan=5 align=\"left\"><b>Service Delivery Tracking for: " + varStartDatePr.Year + "</b></td></tr>");


                            sbHeaderList.AppendLine(clientTotals);

                            sbHeaderList.AppendLine("</table>");
                        }
                        else if (varReport == "KPI - Mallee")
                        {
                            // Get the report names and file names
                            varReportName = "KPI Mallee - " + varStartDatePr.ToString("MMM") + " " + varStartDatePr.Year + " to " + varEndDatePr.ToString("MMM") + " " + varEndDatePr.Year;
                            varFileName = "KPI Mallee - " + varStartDatePr.ToString("MMM") + " " + varStartDatePr.Year + " to " + varEndDatePr.ToString("MMM") + " " + varEndDatePr.Year + ".xls";

                            KPIMalleeString = "<tr><td colspan='4'><b>Support Period Stats</b></td></tr><tr><td><b>Program Name</b></td><td><b>New</b></td><td><b>Active</b></td><td><b>Closed</b></td></tr>";

                            for (var x = 0; x < programsKPIMallee.Length; x++)
                            {
                                KPIMalleeString += "<tr><td><b>" + programNamesKPIMallee[x] + "</b></td>";

                                //Initializing variables each run through
                                aggregateSupport = 0;
                                newSupport = 0;
                                closedSupport = 0;
                                activitiesSum = 0;

                                //Get the support periods created within the selected period of time. - Would have been made during the month.
                                dbSupportList = @"
                                            <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                                <entity name='new_palmclientsupport'>
                                                <attribute name='new_palmclientsupportid' alias='supportNew_count' aggregate='count' />
                                                <filter type='and'>
                                                    <condition attribute='new_startdate' operator='on-or-after' value='" + varStartDate + @"' />
                                                    <condition attribute='new_startdate' operator='on-or-before' value='" + varEndDate + @"' />
                                                    <condition attribute='new_program' operator='eq' uiname='" + programsKPIMallee[x].recName + "' uitype='" + programsKPIMallee[x].theType + "' value='" + programsKPIMallee[x].guid + @"' />
                                                </filter>
                                                </entity>
                                            </fetch>"; // NEW
                                dbSupportList2 = @"
                                            <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                                <entity name='new_palmclientsupport'>
                                                <attribute name='new_palmclientsupportid' alias='supportActive_count' aggregate='count' />
                                                <filter type='and'>
                                                    <condition attribute='new_startdate' operator='on-or-before' value='" + varEndDate + @"' />
                                                    <condition attribute='new_program' operator='eq' uiname='" + programsKPIMallee[x].recName + "' uitype='" + programsKPIMallee[x].theType + "' value='" + programsKPIMallee[x].guid + @"' />
                                                    <filter type='or'>
                                                        <condition attribute='new_enddate' operator='on-or-after' value='" + varStartDate + @"' />
                                                        <condition attribute='new_enddate' operator='null' />
                                                    </filter>
                                                </filter>
                                                </entity>
                                            </fetch>"; //Active
                                dbSupportList3 = @"
                                            <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                                <entity name='new_palmclientsupport'>
                                                <attribute name='new_palmclientsupportid' alias='supportClosed_count' aggregate='count' />
                                                <filter type='and'>
                                                    <condition attribute='new_program' operator='eq' uiname='" + programsKPIMallee[x].recName + "' uitype='" + programsKPIMallee[x].theType + "' value='" + programsKPIMallee[x].guid + @"' />
                                                    <condition attribute='new_enddate' operator='on-or-after' value='" + varStartDate + @"' />
                                                    <condition attribute='new_enddate' operator='on-or-before' value='" + varEndDate + @"' />
                                                </filter>
                                                </entity>
                                            </fetch>"; //CLOSED
                                dbReferralsList = @"
                                            <fetch aggregate='true'>
                                                <entity name='new_palmclientsupport'>
                                                <attribute name='new_palmclientsupportid' alias='supportRef_count' aggregate='count' />
                                                <attribute name='new_reftype' alias='referralType' groupby='true' />
                                                <filter type='and'>
                                                    <condition attribute='new_program' operator='eq' uiname='" + programsKPIMallee[x].recName + "' uitype='" + programsKPIMallee[x].theType + "' value='" + programsKPIMallee[x].guid + @"' />
                                                    <condition attribute='new_enddate' operator='on-or-after' value='" + varStartDate + @"' />
                                                    <condition attribute='new_enddate' operator='on-or-before' value='" + varEndDate + @"' />
                                                    <condition attribute='new_reftype' operator='not-null' />
                                                </filter>
                                                </entity>
                                            </fetch>"; //Closed & Referred - counting on group by.
                                dbActivitiesList = @"
                                            <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                                <entity name='new_palmclientactivities'>
                                                <attribute name='new_amount' alias='activities_sum' aggregate='sum' />
                                                <link-entity name='new_palmclientsupport' from='new_palmclientsupportid' to='new_supportperiod' link-type='outer' />
                                                <filter type='and'>
                                                    <condition attribute='new_entrydate' operator='on-or-before' value='" + varEndDate + @"' />
                                                    <condition attribute='new_entrydate' operator='on-or-after' value='" + varStartDate + @"' />
                                                    <condition entityname='new_palmclientsupport' attribute='new_program' operator='eq' uiname='" + programsKPIMallee[0].recName + "' uitype='" + programsKPIMallee[0].theType + "' value='" + programsKPIMallee[0].guid + @"' />
                                                </filter>
                                                </entity>
                                            </fetch>"; //Activities

                                

                                // Get the fetch XML data and place in entity collection objects
                                result = _service.RetrieveMultiple(new FetchExpression(dbSupportList)); //New
                                result2 = _service.RetrieveMultiple(new FetchExpression(dbSupportList2)); //Active
                                result3 = _service.RetrieveMultiple(new FetchExpression(dbSupportList3)); //Closed
                                result4 = _service.RetrieveMultiple(new FetchExpression(dbReferralsList)); //Referrals
                                result5 = _service.RetrieveMultiple(new FetchExpression(dbActivitiesList)); //Activities

                                // Support Period Numbers
                                foreach (var c in result.Entities)
                                {
                                    aggregateSupport += (Int32)((AliasedValue)c["supportNew_count"]).Value; //New
                                    totalNewSupport += aggregateSupport;
                                }

                                KPIMalleeString += "<td>" + aggregateSupport + "</td>";

                                foreach(var c in result2.Entities)
                                {
                                    newSupport += (Int32)((AliasedValue)c["supportActive_count"]).Value; //Active
                                    totalActiveSupport += newSupport;
                                }

                                KPIMalleeString += "<td>" + newSupport + "</td>";

                                foreach(var c in result3.Entities)
                                {
                                    closedSupport += (Int32)((AliasedValue)c["supportClosed_count"]).Value; //Closed
                                    totalClosedSupport += closedSupport;
                                }

                                foreach(var c in result4.Entities)
                                {
                                    if (c.FormattedValues["referralType"] == "External"){
                                        //external
                                        refCountEx += (Int32)((AliasedValue)c["supportRef_count"]).Value;
                                    }
                                    else if(c.FormattedValues["referralType"] == "Internal"){
                                        refCountIn += (Int32)((AliasedValue)c["supportRef_count"]).Value;
                                    }
                                }
                                

                                KPIMalleeString += "<td>" + closedSupport + "</td>";

                                

                                foreach (var c in result5.Entities)
                                {
                                    activitiesSum += Convert.ToDouble(((AliasedValue)c["activities_sum"]).Value); //Activites
                                }


                            }
                            refCount += refCountEx + refCountIn;
                            KPIMalleeActivities = "<tr><td>&nbsp;</td><td><b>Activity Time</b></td></tr><tr><td>Mallee IAP</td><td>" + activitiesSum + "</td></tr>";
                            KPIMalleeActivities += "<tr><td>&nbsp;</td><td><b>External</b></td><td><b>Internal</b></td><td><b>Referrals Total</b></td></tr><tr><td>Mallee Programs</td><td>" + refCountEx + "</td><td>" + refCountIn + "</td><td>" + refCount + "</td></tr>";

                            KPIMalleeString += "</tr>";
                            KPIMalleeString += "<tr><td><b>TOTALS</b></td><td>" + totalNewSupport + "</td><td>" + totalActiveSupport + "</td><td>" + totalClosedSupport + "</td></tr>";


                            KPIMalleeFinancials = "<tr><td><b>Financial Assistance</b></td></tr>" +
                                "<tr><td>&nbsp;</td><td colspan='2'>Emergency Accommodation</td><td colspan='2'>Rent In Arrears</td><td colspan='2'>Rent In Advance</td><td colspan='2'>Bonds</td><td colspan='2'>Material Aid</td></tr>" +
                                "<tr><td>&nbsp;</td><td>#</td><td>$</td><td>#</td><td>$</td><td>#</td><td>$</td><td>#</td><td>$</td><td>#</td><td>$</td></tr>";
                            for (var j = 0; j < financialKPIMallee.Length; j++)
                            {

                                KPIMalleeFinancials += "<tr><td><b>" + financialNamesKPIMallee[j] + "</b></td>";

                                for (var k = 0; k < financialFilterKPIMallee.Length; k++) {
                                    aggregateFin = 0;
                                    sumFin = 0;

                                    dbFinList = @"
                                        <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                            <entity name='new_palmclientfinancial'>
                                                <attribute name='new_palmclientfinancialid' alias='fin_count' aggregate='count' />
                                                <attribute name='new_amount' alias='fin_sum' aggregate='sum' />
                                                <link-entity name='new_palmsuassistance' from='new_palmsuassistanceid' to='new_assistance' link-type='outer'>
                                                    <attribute name='new_assistance' alias='groupedAssistance' groupby='true' />
                                                </link-entity>
                                                <filter type='and'>
                                                    <condition attribute='new_entrydate' operator='on-or-after' value='" + varStartDate + @"' />
                                                    <condition attribute='new_entrydate' operator='on-or-before' value='" + varEndDate + @"' />
                                                    <condition attribute='new_brokerage' operator='eq' uiname='" + financialKPIMallee[j].recName + "' uitype='" + financialKPIMallee[j].theType + "' value='" + financialKPIMallee[j].guid + @"' />
                                                       " + financialFilterKPIMallee[k] + @"'
                                                </filter>
                                            </entity>
                                        </fetch>"; //SUM & Count of financials.

                                    result = _service.RetrieveMultiple(new FetchExpression(dbFinList));

                                    foreach(var c in result.Entities)
                                    {
                                        if((string)((AliasedValue)c["groupedAssistance"]).Value == "Accommodation")
                                        {
                                            countEmergency += (Int32)((AliasedValue)c["fin_count"]).Value;
                                            sumEmergency += Convert.ToDouble(((AliasedValue)c["fin_sum"]).Value);
                                        }
                                        else if((string)((AliasedValue)c["groupedAssistance"]).Value == "Rent Arrears")
                                        {
                                            countRentArrears += (Int32)((AliasedValue)c["fin_count"]).Value;
                                            sumRentArrears += Convert.ToDouble(((AliasedValue)c["fin_sum"]).Value);
                                        }
                                        else if((string)((AliasedValue)c["groupedAssistance"]).Value == "Rent In Advance")
                                        {
                                            countRentAdvance += (Int32)((AliasedValue)c["fin_count"]).Value;
                                            sumRentAdvance += Convert.ToDouble(((AliasedValue)c["fin_sum"]).Value);
                                        }
                                        else if((string)((AliasedValue)c["groupedAssistance"]).Value == "Bonds")
                                        {
                                            countBonds += (Int32)((AliasedValue)c["fin_count"]).Value;
                                            sumBonds += Convert.ToDouble(((AliasedValue)c["fin_sum"]).Value);
                                        }
                                        else if((string)((AliasedValue)c["groupedAssistance"]).Value == "Aids/Equipment" || (string)((AliasedValue)c["groupedAssistance"]).Value == "Food" || (string)((AliasedValue)c["groupedAssistance"]).Value == "Miscellaneous" || (string)((AliasedValue)c["groupedAssistance"]).Value == "Travel")
                                        {
                                            countMaterial += (Int32)((AliasedValue)c["fin_count"]).Value;
                                            sumMaterial += Convert.ToDouble(((AliasedValue)c["fin_sum"]).Value);
                                        }
                                        aggregateFin += (Int32)((AliasedValue)c["fin_count"]).Value;
                                        sumFin += Convert.ToDouble(((AliasedValue)c["fin_sum"]).Value);
                                    }

                                    if (aggregateFin == 0)
                                        KPIMalleeFinancials += "<td>&nbsp;</td>";
                                    else
                                        KPIMalleeFinancials += "<td>" + aggregateFin + "</td>";

                                    if (sumFin == 0)
                                        KPIMalleeFinancials += "<td>&nbsp;</td>";
                                    else
                                        KPIMalleeFinancials += "<td>$" + sumFin.ToString("#.##") + "</td>"; 

                                }


                                KPIMalleeFinancials += "</tr>";

                            }

                            KPIMalleeFinancials += "<tr><td><b>Totals</b></td><td>" + countEmergency + "</td><td>$" + sumEmergency.ToString("#.##") + "</td>" +
    "<td>" + countRentArrears + "</td><td>$" + sumRentArrears.ToString("#.##") + "</td>" +
    "<td>" + countRentAdvance + "</td><td>$" + sumRentAdvance.ToString("#.##") + "</td>" +
    "<td>" + countBonds + "</td><td>$" + sumBonds.ToString("#.##") + "</td>" +
    "<td>" + countMaterial + "</td><td>$" + sumMaterial.ToString("#.##") + "</td></tr>";


                            //Header part of the summary report
                            sbHeaderList.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                            sbHeaderList.AppendLine("<head>");
                            sbHeaderList.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                            sbHeaderList.AppendLine("<!--[if gte mso 9]>");
                            sbHeaderList.AppendLine("<xml>");
                            sbHeaderList.AppendLine("<x:ExcelWorkbook>");
                            sbHeaderList.AppendLine("<x:ExcelWorksheets>");
                            sbHeaderList.AppendLine("<x:ExcelWorksheet>");

                            //this line names the worksheet
                            sbHeaderList.AppendLine("<x:Name>Report Data</x:Name>");

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

                            // Produce the summary using the values above
                            sbHeaderList.AppendLine("<table border=0 cellpadding=3>");

                            sbHeaderList.AppendLine(KPIMalleeString);
                            sbHeaderList.AppendLine("<tr><td>&nbsp;</td></tr>");
                            sbHeaderList.AppendLine(KPIMalleeActivities);
                            sbHeaderList.AppendLine("<tr><td>&nbsp;</td></tr>");
                            sbHeaderList.AppendLine(KPIMalleeFinancials);

                            sbHeaderList.AppendLine("</table>");

                        }

                        byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                        string encodedData = System.Convert.ToBase64String(filename);
                        Entity Annotation = new Entity("annotation");
                        Annotation.Attributes["objectid"] = new EntityReference("new_connectgoreport", varReportId);
                        Annotation.Attributes["objecttypecode"] = "new_connectgoreport";
                        Annotation.Attributes["subject"] = "Report";
                        Annotation.Attributes["documentbody"] = encodedData;
                        Annotation.Attributes["mimetype"] = @"application/msexcel";
                        Annotation.Attributes["notetext"] = varReportName;
                        Annotation.Attributes["filename"] = varFileName;
                        _service.Create(Annotation);

                        sbHeaderList.Clear();
                        sbReportList.Clear();

                    }


                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in the Report plug-in." + ex, ex);
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

        // Clean date in DEX format for year only
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

        //Return start of week (Monday) for the date passed
        public DateTime weekStart(DateTime getdate)
        {
            if (getdate.ToString("dddd") == "Tuesday")
                getdate = getdate.AddDays(-1);
            else if (getdate.ToString("dddd") == "Wednesday")
                getdate = getdate.AddDays(-2);
            else if (getdate.ToString("dddd") == "Thursday")
                getdate = getdate.AddDays(-3);
            else if (getdate.ToString("dddd") == "Friday")
                getdate = getdate.AddDays(-4);
            else if (getdate.ToString("dddd") == "Saturday")
                getdate = getdate.AddDays(-5);
            else if (getdate.ToString("dddd") == "Sunday")
                getdate = getdate.AddDays(-6);

            return getdate;
        }
    }

    public class Record
    {
        public string guid;
        public string recName;
        public string theType;

        public Record(string id, string name, string type)
        {
            guid = id;
            recName = name;
            theType = type;
        }
    }
}



