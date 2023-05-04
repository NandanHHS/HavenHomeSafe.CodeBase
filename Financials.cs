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
    public class Financials : IPlugin
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

                // Verify that the target entity represents the Palm Go MDS entity
                // If not, this plug-in was not registered correctly.
                // if (entity.LogicalName != "new_palmgofinancial")
                //    return;

                try
                {
                    // Fixes:
                    // Remove ampersand from comments
                    // Compare values to Lower

                    string varBrokerageId = ""; // Brokerage Id
                    string varBrokerage = ""; // Brokerage Description
                    string varDescription = ""; // Report Description
                    string varReport = ""; // Report Type
                    bool varCreateExtract = false; // Whether to create report
                    OptionSetValue varBankAccount = new OptionSetValue(); // Bank Account selected
                    Guid varFinancialID = new Guid(); // GUID for financial report
                    StringBuilder sbHeaderList = new StringBuilder(); // String builder for header
                    StringBuilder sbReportList = new StringBuilder(); // String builder for report
                    StringBuilder sbReport2List = new StringBuilder(); // String builder for report section
                    StringBuilder sbFinancialList = new StringBuilder(); // String builder for financial list
                    StringBuilder sbErrorList = new StringBuilder(); // String builder for errors

                    string varFileName = "No Report"; // Report file name
                    string varFileName2 = "No Report"; // Report file name
                    DateTime varStartDate = new DateTime(); // Report start date
                    DateTime varEndDate = new DateTime(); // Report end date
                    int varCheckInt = 0; // Variable to parse integers
                    double varCheckDouble = 0; // Variable to parse doubles
                    DateTime varCheckDate = new DateTime(); // Variable to parse dates
                    EntityReference getEntity; // Object for entity reference
                    AliasedValue getAlias; // Object for aliased value
                    DateTime varWeekStart = new DateTime(); //Start of week
                    DateTime varWeekEnd = new DateTime(); //End of week
                    DateTime varMonthStart = new DateTime(); //Start of month
                    DateTime varMonthEnd = new DateTime(); //End of month
                    DateTime varYearStart = new DateTime(); //Start of year
                    DateTime varYearEnd = new DateTime(); //End of year
                    string varYear = ""; // Year part
                    TimeSpan ts1 = new TimeSpan(); // Timespan object for dates
                    TimeSpan ts2 = new TimeSpan(); // Timespan object for dates
                    string varReportName = "No Report"; // Name of report
                    string varReportName2 = "No Report"; // Name of report
                    string varPeriod = "month"; // Report period

                    // Variables for data
                    string dbFinancialList = "";
                    string dbFinancialList2 = ""; //Extra created for purpose of nonclient fins in Kypera Spend Report.
                    string dbSpendList = "";
                    string dbBudgetList = "";
                    string dbEpisodeList = "";
                    string dbWeekAmList = "";
                    string dbMonthAmList = "";
                    string dbYearAmList = "";
                    string dbYTDWAmList = "";
                    string dbYTDMAmList = "";
                    string varWeekPr = ""; //Print week period
                    string varMonthPr = ""; //Print month period
                    string varYearPr = ""; //Print year period

                    string dbPalmClientFinancialId = "";
                    string dbEntryDate = "";
                    string dbPaidDate = "";
                    string dbCheque = "";
                    string dbVoucher = "";
                    string dbPurpose = "";
                    string dbPayee = "";
                    double dbAmount = 0;
                    double dbGst = 0;
                    string dbGstCode = "";
                    string dbFirstNameNc = "";
                    string dbSurnameNc = "";
                    string dbFirstName = "";
                    string dbSurname = "";
                    string dbAssistance = "";
                    string dbPaidBy = "";
                    double varTotal = 0;
                    double dbYtdSpend = 0;
                    double varYTDB = 0;
                    double varAmount = 0;
                    string dbPurchaseType = "";
                    string varNMDStype = "";

                    string dbBankAcc = "";
                    string dbProgId = "";

                    string dbPalmClientId = "";
                    string dbClient = "";
                    string dbGender = "";
                    string dbIndigenous = "";
                    string dbAge = ""; //Variable for age of client.
                    string dbDob = "";
                    string dbDobEst = "";
                    string dbDobEstD = ""; // DOB day estimated field
                    string dbDobEstM = ""; // DOB month estimated field
                    string dbDobEstY = ""; // DOB year estimated field
                    string dbMdsSlk = "";
                    string dbShorSlk = "";
                    string dbWeekRent = "";
                    string dbPalmClientSupportId = "";
                    string dbPalmClientSupportIdOld = ""; //Support Period id - old
                    string dbDoShor = "";
                    string dbShorAgency = "";
                    string dbLocality = "";
                    string dbState = "";
                    string dbPostcode = "";
                    string dbPuhId = "";
                    string dbBrokerage = "";

                    string dbPuhSupportPeriod = "";
                    string dbPuhIdOld = ""; //Presenting unit head old
                    string dbPuhClient = "";

                    //Variable for setting old ids for imported records.
                    string varPrintSpId = "";
                    string varPrintPuhId = "";

                    string dbAssistanceCode = "";
                    string dbGeneratedId = "";

                    string dbAccountNum = "";
                    string varGSTUpDown = "";

                    // Variables for SLK
                    string varSLK = "";
                    string varSurname = "";
                    string varFirstName = "";
                    string varDob = "";
                    string varGender = "";


                    string varDobFlag = ""; // Dob flag

                    double varWeekAm = 0; //Amount for week
                    double varMonthAm = 0; //Amount for month
                    double varYearAm = 0; //Amount for year
                    double varYTDWAm = 0; //YTD amount for week
                    double varYTDMAm = 0; //YTD amount for month
                    double varDayFig = 0; //Figure for the day
                    double varWeekTot = 0; //Total for the week
                    double varMonthTot = 0; //Total for the month
                    double varYearTot = 0; //Total for the year
                    double varYTDWTot = 0; //Total YTD for the week
                    double varYTDMTot = 0; //Total YTD for the month
                    int varWeek = 1; //Counter to loop through the weeks
                    int varRowId = 1; //Ids for table rows
                    double varTotalPr = 0; //Total for period
                    int varEpisode = 0; //Episode count
                    double varStartFig = 0; //Starting figure for the period
                    int varCount = 0; // Counter

                    string varHex = "99"; // Row colour
                    string varCol = ""; //Row colour
                    bool varReport2 = false; // Whether second report produced

                    // Entity collection objects for fetch XML data
                    EntityCollection result;
                    EntityCollection result2;
                    EntityCollection result3;
                    EntityCollection result4;
                    EntityCollection result5;
                    EntityCollection result6;
                    EntityCollection result7;
                    EntityCollection result8;

                    string varTest = ""; // Debug

                    // Only do this if the entity is the Palm Go Financials entity
                    if (entity.LogicalName == "new_palmgofinancials")
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
                        varCreateExtract = entity.GetAttributeValue<bool>("new_createextract");
                        varBankAccount = entity.GetAttributeValue<OptionSetValue>("new_bankacount");
                        varStartDate = entity.GetAttributeValue<DateTime>("new_datefrom");
                        varEndDate = entity.GetAttributeValue<DateTime>("new_dateto");


                        //Get the start and end of month
                        varMonthStart = Convert.ToDateTime("1-" + varStartDate.ToString("MMM") + "-" + varStartDate.Year);
                        varMonthEnd = varMonthStart.AddMonths(1);
                        varMonthEnd = varMonthEnd.AddDays(-1);

                        //Get the week start and end dates
                        varWeekStart = weekStart(varMonthStart); //Start of the week must be a Monday
                        varWeekEnd = varWeekStart.AddDays(6);

                        //Get the start and end of financial year based on the month
                        if (varMonthStart.ToString("MMM") == "Jan" || varMonthStart.ToString("MMM") == "Feb" || varMonthStart.ToString("MMM") == "Mar" || varMonthStart.ToString("MMM") == "Apr" || varMonthStart.ToString("MMM") == "May" || varMonthStart.ToString("MMM") == "Jun")
                        {
                            varYearStart = Convert.ToDateTime("1-Jul-" + (Convert.ToInt32(varMonthStart.Year) - 1));
                            varYearEnd = Convert.ToDateTime("30-Jun-" + varMonthStart.Year);
                            varYear = (varMonthStart.Year - 1).ToString();
                        }
                        else
                        {
                            varYearStart = Convert.ToDateTime("1-Jul-" + varMonthStart.Year);
                            varYearEnd = Convert.ToDateTime("30-Jun-" + (Convert.ToInt32(varMonthStart.Year) + 1));
                            varYear = varMonthStart.Year.ToString();
                        }

                        // Get the report type
                        if (entity.Contains("new_report"))
                            varReport = entity.FormattedValues["new_report"];

                        //if (entity.Contains("new_bankaccount"))
                            //varBankAccount = (OptionSetValue)entity.Attributes["new_bankaccount"];



                        // Get the entity ID
                        varFinancialID = entity.Id;

                        // Get the associated data from the brokerage chosen
                        EntityReference ownerLookup = (EntityReference)entity.Attributes["new_brokerage"];
                        varBrokerageId += ownerLookup.Id.ToString() + ".\r\n";
                        varBrokerageId += ((EntityReference)entity.Attributes["new_brokerage"]).Name + ".\r\n";
                        varBrokerageId += ownerLookup.LogicalName + ".\r\n";

                        var actualOwningUnit = _service.Retrieve(ownerLookup.LogicalName, ownerLookup.Id, new ColumnSet(true));

                        // Get brokerage short hand
                        varBrokerage = actualOwningUnit["new_shorthand"].ToString();

                        // Get the report period or set a default
                        if (varPeriod != "month" && varPeriod != "year")
                            varPeriod = "week";

                        //Get the printable periods
                        varWeekPr = varWeekStart.Day + "/" + varWeekStart.Month + " to " + varWeekEnd.Day + "/" + varWeekEnd.Month;
                        varMonthPr = varMonthStart.ToString("MMMM");
                        varYearPr = varYearStart.Year + " - " + varYearEnd.Year;

                        // This data is processed for thebudget report
                        if (varReport == "Budget")
                        {
                            varReport2 = true; // Two reports are produced

                            // Get the report names and file names
                            varReportName = varBrokerage + " budget report for " + varYearStart.ToString("MMM") + " " + varYearStart.Year + " to " + varYearEnd.ToString("MMM") + " " + varYearEnd.Year;
                            varReportName2 = varBrokerage + " financial data for " + varYearStart.ToString("MMM") + " " + varYearStart.Year + " to " + varYearEnd.ToString("MMM") + " " + varYearEnd.Year;
                            varFileName = varBrokerage + " budget report for " + varYearStart.ToString("MMM") + " " + varYearStart.Year + " to " + varYearEnd.ToString("MMM") + " " + varYearEnd.Year + ".xls";
                            varFileName2 = varBrokerage + " financial for " + varYearStart.ToString("MMM") + " " + varYearStart.Year + " to " + varYearEnd.ToString("MMM") + " " + varYearEnd.Year + ".xls";

                            //Get the financials for the episodes count
                            dbEpisodeList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclientfinancial'>
                                <attribute name='new_palmclientfinancialid' />
                                <attribute name='new_entrydate' />
                                <attribute name='new_datepaid' />
                                <attribute name='new_cheque' />
                                <attribute name='new_purpose' />
                                <attribute name='new_payee' />
                                <attribute name='new_amount' />
                                <attribute name='new_gst' />
                                <attribute name='new_firstnamenc' />
                                <attribute name='new_surnamenc' />
                                <attribute name='new_firstname' />
                                <attribute name='new_surname' />
                                <attribute name='new_paidby' />
                                <attribute name='new_assistance' />
                                <filter type='and'>
                                    <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + varYearStart + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='le' value='" + varYearEnd + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_paidby' operator='ne' value='100000003' />
                                </filter >
                                <order entityname='new_palmclientfinancial' attribute='new_entrydate' />
                              </entity>
                            </fetch> ";

                            //Get the budget figure for the year based on the brokerage type/s chosen
                            dbBudgetList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                                <entity name='new_palmsubudget'>
                                <attribute name='new_amount'/>
                                <attribute name='new_startdate'/>
                                <attribute name='new_enddate'/>
                                <attribute name='new_palmsubudgetid'/>
                                <filter type='and'>
                                    <condition entityname='new_palmsubudget' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmsubudget' attribute='new_startdate' operator='le' value='" + varYearStart + @"' />
                                    <condition entityname='new_palmsubudget' attribute='new_enddate' operator='ge' value='" + varYearStart + @"' />
                                </filter>
                                </entity>
                            </fetch> ";

                            //Get the total amount elapsed for the week
                            dbWeekAmList = @"
                                <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                    <entity name='new_palmclientfinancial'>
                                    <attribute name='new_amount' alias='totalamount_sum' aggregate='sum'/> 
                                    <filter type='and'>
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + varWeekStart + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='lt' value='" + varWeekEnd + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_paidby' operator='ne' value='100000003' />
                                    </filter>
                                    </entity>
                                </fetch> ";

                            //Get the total amount elapsed for the month
                            dbMonthAmList = @"
                                <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                    <entity name='new_palmclientfinancial'>
                                    <attribute name='new_amount' alias='totalamount_sum' aggregate='sum'/> 
                                    <filter type='and'>
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + varMonthStart + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='lt' value='" + varMonthEnd + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_paidby' operator='ne' value='100000003' />
                                    </filter>
                                    </entity>
                                </fetch> ";

                            //Get the total amount elapsed for the year
                            dbYearAmList = @"
                                <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                    <entity name='new_palmclientfinancial'>
                                    <attribute name='new_amount' alias='totalamount_sum' aggregate='sum'/> 
                                    <filter type='and'>
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + varYearStart + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='lt' value='" + varYearEnd + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_paidby' operator='ne' value='100000003' />
                                    </filter>
                                    </entity>
                                </fetch> ";

                            //Get the YTD week amount elapsed
                            dbYTDWAmList = @"
                                <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                    <entity name='new_palmclientfinancial'>
                                    <attribute name='new_amount' alias='totalamount_sum' aggregate='sum'/> 
                                    <filter type='and'>
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + varYearStart + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='lt' value='" + varWeekEnd + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_paidby' operator='ne' value='100000003' />
                                    </filter>
                                    </entity>
                                </fetch> ";

                            //Get the YTD month amount elapsed
                            dbYTDMAmList = @"
                                <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                    <entity name='new_palmclientfinancial'>
                                    <attribute name='new_amount' alias='totalamount_sum' aggregate='sum'/> 
                                    <filter type='and'>
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + varYearStart + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='lt' value='" + varMonthEnd + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_paidby' operator='ne' value='100000003' />
                                    </filter>
                                    </entity>
                                </fetch> ";

                            // Get the fetch XML data and place in entity collection objects
                            result = _service.RetrieveMultiple(new FetchExpression(dbEpisodeList));
                            result3 = _service.RetrieveMultiple(new FetchExpression(dbBudgetList));
                            result4 = _service.RetrieveMultiple(new FetchExpression(dbWeekAmList));
                            result5 = _service.RetrieveMultiple(new FetchExpression(dbMonthAmList));
                            result6 = _service.RetrieveMultiple(new FetchExpression(dbYearAmList));
                            result7 = _service.RetrieveMultiple(new FetchExpression(dbYTDWAmList));
                            result8 = _service.RetrieveMultiple(new FetchExpression(dbYTDMAmList));

                            // Budget figure
                            foreach (var c in result3.Entities)
                            {
                                if (c.FormattedValues.Contains("new_amount"))
                                    Double.TryParse(c.FormattedValues["new_amount"], out varAmount);
                                else if (c.Attributes.Contains("new_amount"))
                                    Double.TryParse(c.Attributes["new_amount"].ToString(), out varAmount);
                                else
                                    varAmount = 0;
                            } // spend loop

                            // Weekly Spend
                            foreach (var c in result4.Entities)
                            {
                                if (c.FormattedValues.Contains("totalamount_sum"))
                                    Double.TryParse(c.FormattedValues["totalamount_sum"], out varWeekAm);
                                else if (c.Attributes.Contains("totalamount_sum"))
                                    Double.TryParse(c.Attributes["totalamount_sum"].ToString(), out varWeekAm);
                                else
                                    varWeekAm = 0;
                            } // spend loop

                            // Monthly Spend
                            foreach (var c in result5.Entities)
                            {
                                if (c.FormattedValues.Contains("totalamount_sum"))
                                    Double.TryParse(c.FormattedValues["totalamount_sum"], out varMonthAm);
                                else if (c.Attributes.Contains("totalamount_sum"))
                                    Double.TryParse(c.Attributes["totalamount_sum"].ToString(), out varMonthAm);
                                else
                                    varMonthAm = 0;
                            } // spend loop

                            // Annual Spend
                            foreach (var c in result6.Entities)
                            {
                                if (c.FormattedValues.Contains("totalamount_sum"))
                                    Double.TryParse(c.FormattedValues["totalamount_sum"], out varYearAm);
                                else if (c.Attributes.Contains("totalamount_sum"))
                                    Double.TryParse(c.Attributes["totalamount_sum"].ToString(), out varYearAm);
                                else
                                    varYearAm = 0;
                            } // spend loop

                            // YTD weekly amount
                            foreach (var c in result7.Entities)
                            {
                                if (c.FormattedValues.Contains("totalamount_sum"))
                                    Double.TryParse(c.FormattedValues["totalamount_sum"], out varYTDWAm);
                                else if (c.Attributes.Contains("totalamount_sum"))
                                    Double.TryParse(c.Attributes["totalamount_sum"].ToString(), out varYTDWAm);
                                else
                                    varYTDWAm = 0;
                            } // spend loop

                            // YTD monthly amount
                            foreach (var c in result8.Entities)
                            {
                                if (c.FormattedValues.Contains("totalamount_sum"))
                                    Double.TryParse(c.FormattedValues["totalamount_sum"], out varYTDMAm);
                                else if (c.Attributes.Contains("totalamount_sum"))
                                    Double.TryParse(c.Attributes["totalamount_sum"].ToString(), out varYTDMAm);
                                else
                                    varYTDMAm = 0;
                            } // spend loop

                            //Get the daily budget figure
                            ts1 = (varYearEnd - varYearStart);
                            varDayFig = (double)varAmount / (ts1.TotalDays + 1);

                            //Total budget for the week
                            ts1 = (varWeekEnd - varWeekStart);
                            varWeekTot = (double)(ts1.TotalDays + 1) * varDayFig;
                            //Total budget for the month
                            ts1 = (varMonthEnd - varMonthStart);
                            varMonthTot = (double)(ts1.TotalDays + 1) * varDayFig;
                            //Total budget for the year
                            varYearTot = varAmount;
                            //Total budget for the YTD week
                            ts1 = (varWeekEnd - varYearStart);
                            varYTDWTot = (double)(ts1.TotalDays + 1) * varDayFig;
                            //Total budget for the YTD month
                            ts1 = (varMonthEnd - varYearStart);
                            varYTDMTot = (double)(ts1.TotalDays + 1) * varDayFig;


                            //Header part of the data report
                            sbReport2List.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                            sbReport2List.AppendLine("<head>");
                            sbReport2List.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                            sbReport2List.AppendLine("<!--[if gte mso 9]>");
                            sbReport2List.AppendLine("<xml>");
                            sbReport2List.AppendLine("<x:ExcelWorkbook>");
                            sbReport2List.AppendLine("<x:ExcelWorksheets>");
                            sbReport2List.AppendLine("<x:ExcelWorksheet>");

                            //this line names the worksheet
                            sbReport2List.AppendLine("<x:Name>Financial Data</x:Name>");

                            sbReport2List.AppendLine("<x:WorksheetOptions>");

                            sbReport2List.AppendLine("<x:Panes>");
                            sbReport2List.AppendLine("</x:Panes>");
                            sbReport2List.AppendLine("</x:WorksheetOptions>");
                            sbReport2List.AppendLine("</x:ExcelWorksheet>");
                            sbReport2List.AppendLine("</x:ExcelWorksheets>");
                            sbReport2List.AppendLine("</x:ExcelWorkbook>");
                            sbReport2List.AppendLine("</xml>");
                            sbReport2List.AppendLine("<![endif]-->");
                            sbReport2List.AppendLine("</head>");

                            sbReport2List.AppendLine("<table border=0 cellpadding=3>");
                            sbReport2List.AppendLine("<tr>");
                            sbReport2List.AppendLine("<td class=\"prBorderNW\">Entry Date</td>");
                            sbReport2List.AppendLine("<td class=\"prBorderNW\">Date Paid</td>");
                            sbReport2List.AppendLine("<td class=\"prBorderNW\">Assistance</td>");
                            sbReport2List.AppendLine("<td class=\"prBorderNW\">Paid By</td>");
                            sbReport2List.AppendLine("<td class=\"prBorderNW\">Cheque Number</td>");
                            sbReport2List.AppendLine("<td class=\"prBorder\">Client Name</td>");
                            sbReport2List.AppendLine("<td class=\"prBorderNW\">Purpose</td>");
                            sbReport2List.AppendLine("<td class=\"prBorder\">Payee</td>");
                            sbReport2List.AppendLine("<td class=\"prBorderNW\">Amount</td>");
                            sbReport2List.AppendLine("<td class=\"prBorderNW\">GST</td>");
                            sbReport2List.AppendLine("</tr>");

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
                            sbHeaderList.AppendLine("<x:Name>Financial Data</x:Name>");

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
                            sbHeaderList.AppendLine("<tr><td colspan=5 align=\"left\"><b>Summary of funds for " + varYearPr + "</b></td></tr>");
                            sbHeaderList.AppendLine("<tr><td colspan=5 align=\"left\"><b>Budget:</b> " + varAmount.ToString("C") + "</td></tr>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td>&nbsp;</td>");
                            sbHeaderList.AppendLine("<td>Spent</td>");
                            sbHeaderList.AppendLine("<td>Allocated</td>");
                            sbHeaderList.AppendLine("<td>Variance</td>");
                            sbHeaderList.AppendLine("<td>YTD Spent</td>");
                            sbHeaderList.AppendLine("<td>YTD Variance</td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td>" + (varWeekPr) + ":</td>");
                            sbHeaderList.AppendLine("<td>" + ((varWeekAm).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("<td>" + ((varWeekTot).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("<td>" + ((varWeekTot - varWeekAm).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("<td>" + ((varYTDWAm).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("<td>" + ((varYTDWTot - varYTDWAm).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td>" + (varMonthPr) + ":</td>");
                            sbHeaderList.AppendLine("<td>" + ((varMonthAm).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("<td>" + ((varMonthTot).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("<td>" + ((varMonthTot - varMonthAm).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("<td>" + ((varYTDMAm).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("<td>" + ((varYTDMTot - varYTDMAm).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td>" + (varYearPr) + ":</td>");
                            sbHeaderList.AppendLine("<td>" + ((varYearAm).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("<td>" + ((varYearTot).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("<td>" + ((varYearTot - varYearAm).ToString("C")) + "</td>");
                            sbHeaderList.AppendLine("<td>&nbsp;</td>");
                            sbHeaderList.AppendLine("<td>&nbsp;</td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.Append("<tr><td colspan=4>");

                            //Determine amount remaining or if there is no money remaining
                            if ((varYearTot - varYearAm) > 0)
                            {
                                ts1 = varYearEnd - DateTime.Now;

                                sbHeaderList.Append("<b>" + (varYearTot - varYearAm).ToString("C") + "</b> remaining over <b>" + (ts1.Days + 1) + "</b> day");

                                if (ts1.Days != 1)
                                    sbHeaderList.Append("s");

                                sbHeaderList.Append(" (<b>" + ((double)(varYearTot - varYearAm) / (ts1.Days + 1)).ToString("C") + "</b> per day)");
                            }
                            else
                                sbHeaderList.Append("There is no money remaining for this account");

                            sbHeaderList.AppendLine("</td></tr>");
                            sbHeaderList.AppendLine("</table>");


                            // Financial Data Loop
                            //varWeekStart and varWeekEnd will be the period variables
                            //Begin from the start of the year (may be a half week)
                            varWeekStart = varYearStart;

                            //The end of the month must be the end of the first month
                            varMonthEnd = varWeekStart.AddMonths(1);
                            varMonthEnd = varMonthEnd.AddDays(-1);

                            //If the period is month or year, set varWeekEnd to the correct end of period
                            if (varPeriod == "month")
                                varWeekEnd = varMonthEnd;
                            else if (varPeriod == "year")
                                varWeekEnd = varYearEnd;

                            //Loop through the data until the end of the year is reached
                            while (varWeekStart < varYearEnd)
                            {
                                //If the period hasn't begun, set the text colour to light grey
                                if (DateTime.Now < varWeekStart)
                                    varHex = "99";

                                //Begin to build the financial row
                                sbFinancialList.AppendLine("<tr style=\"color:#" + varHex + varHex + varHex + ";\" id=\"" + varRowId + "\">");

                                sbFinancialList.AppendLine("<td>&nbsp;</td>");

                                //Enter the period and the date range
                                sbFinancialList.AppendLine("<td>" + varPeriod + " " + varWeek + "</td>\r\n<td>" + cleanDate(varWeekStart) + " - " + cleanDate(varWeekEnd) + "</td>");

                                //Reset the period total and episode count
                                varTotalPr = 0;
                                varEpisode = 0;

                                //Get the period to date budget (allowable spend)
                                ts1 = varWeekEnd - varWeekStart;
                                varStartFig = (double)varStartFig + (varDayFig * (ts1.TotalDays + 1));


                                //Loop through the episodes list, starting from the previous episode
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
                                    if (c.FormattedValues.Contains("new_palmclientfinancialid"))
                                        dbPalmClientFinancialId = c.FormattedValues["new_palmclientfinancialid"];
                                    else if (c.Attributes.Contains("new_palmclientfinancialid"))
                                        dbPalmClientFinancialId = c.Attributes["new_palmclientfinancialid"].ToString();
                                    else
                                        dbPalmClientFinancialId = "";

                                    if (c.FormattedValues.Contains("new_entrydate"))
                                        dbEntryDate = c.FormattedValues["new_entrydate"];
                                    else if (c.Attributes.Contains("new_entrydate"))
                                        dbEntryDate = c.Attributes["new_entrydate"].ToString();
                                    else
                                        dbEntryDate = "";

                                    // Convert date from American format to Australian format
                                    dbEntryDate = cleanDateAM(dbEntryDate);

                                    if (c.FormattedValues.Contains("new_datepaid"))
                                        dbPaidDate = c.FormattedValues["new_datepaid"];
                                    else if (c.Attributes.Contains("new_datepaid"))
                                        dbPaidDate = c.Attributes["new_datepaid"].ToString();
                                    else
                                        dbPaidDate = "";

                                    // Convert date from American format to Australian format
                                    dbPaidDate = cleanDateAM(dbPaidDate);

                                    if (c.FormattedValues.Contains("new_cheque"))
                                        dbCheque = c.FormattedValues["new_cheque"];
                                    else if (c.Attributes.Contains("new_cheque"))
                                        dbCheque = c.Attributes["new_cheque"].ToString();
                                    else
                                        dbCheque = "";



                                    if (c.FormattedValues.Contains("new_purpose"))
                                        dbPurpose = c.FormattedValues["new_purpose"];
                                    else if (c.Attributes.Contains("new_purpose"))
                                        dbPurpose = c.Attributes["new_purpose"].ToString();
                                    else
                                        dbPurpose = "";

                                    if (c.FormattedValues.Contains("new_payee"))
                                        dbPayee = c.FormattedValues["new_payee"];
                                    else if (c.Attributes.Contains("new_payee"))
                                        dbPayee = c.Attributes["new_payee"].ToString();
                                    else
                                        dbPayee = "";

                                    if (c.FormattedValues.Contains("new_amount"))
                                        Double.TryParse(c.FormattedValues["new_amount"], out dbAmount);
                                    else if (c.Attributes.Contains("new_amount"))
                                        Double.TryParse(c.Attributes["new_amount"].ToString(), out dbAmount);
                                    else
                                        dbAmount = 0;

                                    if (c.FormattedValues.Contains("new_gst"))
                                        Double.TryParse(c.FormattedValues["new_gst"], out dbGst);
                                    else if (c.Attributes.Contains("new_gst"))
                                        Double.TryParse(c.Attributes["new_gst"].ToString(), out dbGst);
                                    else
                                        dbGst = 0;

                                    if (c.FormattedValues.Contains("new_firstnamenc"))
                                        dbFirstNameNc = c.FormattedValues["new_firstnamenc"];
                                    else if (c.Attributes.Contains("new_firstnamenc"))
                                        dbFirstNameNc = c.Attributes["new_firstnamenc"].ToString();
                                    else
                                        dbFirstNameNc = "";

                                    if (c.FormattedValues.Contains("new_firstname"))
                                        dbFirstName = c.FormattedValues["new_firstname"];
                                    else if (c.Attributes.Contains("new_firstname"))
                                        dbFirstName = c.Attributes["new_firstname"].ToString();
                                    else
                                        dbFirstName = "";

                                    if (String.IsNullOrEmpty(dbFirstName) == true)
                                        dbFirstName = dbFirstNameNc;

                                    if (c.FormattedValues.Contains("new_surnamenc"))
                                        dbSurnameNc = c.FormattedValues["new_surnamenc"];
                                    else if (c.Attributes.Contains("new_surnamenc"))
                                        dbSurnameNc = c.Attributes["new_surnamenc"].ToString();
                                    else
                                        dbSurnameNc = "";

                                    if (c.FormattedValues.Contains("new_surname"))
                                        dbSurname = c.FormattedValues["new_surname"];
                                    else if (c.Attributes.Contains("new_surname"))
                                        dbSurname = c.Attributes["new_surname"].ToString();
                                    else
                                        dbSurname = "";

                                    if (String.IsNullOrEmpty(dbSurname) == true)
                                        dbSurname = dbSurnameNc;

                                    if (c.FormattedValues.Contains("new_paidby"))
                                        dbPaidBy = c.FormattedValues["new_paidby"];
                                    else if (c.Attributes.Contains("new_paidby"))
                                        dbPaidBy = c.Attributes["new_paidby"].ToString();
                                    else
                                        dbPaidBy = "";

                                    if (c.FormattedValues.Contains("new_assistance"))
                                        dbAssistance = c.FormattedValues["new_assistance"];
                                    else if (c.Attributes.Contains("new_assistance"))
                                        dbAssistance = c.Attributes["new_assistance"].ToString();
                                    else
                                        dbAssistance = "";

                                    //If the entry date is not for this period, break from the loop
                                    if (Convert.ToDateTime(dbEntryDate) > Convert.ToDateTime(varWeekEnd))
                                        break;

                                    if (Convert.ToDateTime(dbEntryDate) >= Convert.ToDateTime(varWeekStart))
                                    {
                                        //Add the amount to the total for the period
                                        varTotalPr += dbAmount;
                                        //Add the amount to the overall total
                                        varTotal += dbAmount;

                                        //Subtract the amount from the allowable spend
                                        varStartFig -= dbAmount;

                                        //Add to the total episodes
                                        varEpisode++;

                                        //Increase the loop starting point
                                        varCount++;

                                        sbReportList.AppendLine("<tr><td colspan=2>&nbsp;</td>\r\n<td>" + varEpisode + "</td>\r\n<td>" + dbEntryDate + "</td>\r\n<td>" + dbAmount.ToString("C") + "</td>\r\n<tr>");
                                        sbReport2List.AppendLine("<tr>\r\n<td class=\"prBorder\">" + dbEntryDate + "</td>\r\n<td class=\"prBorder\">" + dbPaidDate + "</td>\r\n<td class=\"prBorder\">" + dbAssistance + "</td>\r\n<td class=\"prBorder\">" + dbPaidBy + "</td>\r\n<td class=\"prBorder\">" + dbCheque + "</td>\r\n<td class=\"prBorder\">" + dbFirstName + " " + dbSurname + "</td>\r\n<td class=\"prBorder\">" + dbPurpose + "</td>\r\n<td class=\"prBorder\">" + dbPayee + "</td>\r\n<td class=\"prBorder\">" + dbAmount.ToString("C") + "</td>\r\n<td class=\"prBorder\">" + dbGst.ToString("C") + "</td>\r\n</tr>");
                                    }
                                } //Episodes


                                //If the period has not begun, set the total spent to 0
                                if (DateTime.Now < varWeekStart)
                                    varTotal = 0;

                                //Append the episode count to the financial row
                                sbFinancialList.AppendLine("<td colspan=2>&nbsp;</td>\r\n<td>" + varEpisode + "</td>");


                                //If the period has begun determine if the amount spent figure should be red or green
                                if (DateTime.Now >= varWeekStart)
                                {
                                    ts1 = varWeekEnd - varWeekStart;

                                    //If the total spent is less than the amount allocated set to green
                                    if (varTotalPr <= (varDayFig * (ts1.TotalDays + 1)))
                                        varCol = "009900";
                                    //Otherwise set to red
                                    else
                                        varCol = "990000";

                                    sbFinancialList.Append("<td style=\"color: #" + varCol + ";\">");
                                }
                                else
                                    sbFinancialList.Append("<td>");

                                //Append to the financial row: amount spent for the period, amount allocated, and total spent
                                ts1 = varWeekEnd - varWeekStart;
                                sbFinancialList.AppendLine(varTotalPr.ToString("C") + "</td>\r\n<td>(" + ((double)(varDayFig * (ts1.TotalDays + 1)) - varTotalPr).ToString("C") + ")</td>\r\n<td>" + varTotal.ToString("C") + "</td>");

                                //If the period has begun determine if the remaining figure should be red or green
                                if (DateTime.Now >= varWeekStart)
                                {
                                    //If the amount remaining is greater or equal to zero set to green
                                    if (varStartFig >= 0)
                                        varCol = "73D618";
                                    //Otherwise set to red
                                    else
                                        varCol = "F70000";

                                    sbFinancialList.Append("<td style=\"color: #" + varCol + ";\">");
                                }
                                else
                                    sbFinancialList.Append("<td>");

                                //Append to the financial list the amount remaining and the printable report
                                sbFinancialList.AppendLine(varStartFig.ToString("C") + "</td>\r\n<td>&nbsp;</td>\r\n</tr>");

                                //Move to the next row id
                                varRowId++;

                                //Reset the total, episode and category variables
                                varTotalPr = 0;
                                varEpisode = 0;

                                //Get the start date of the next period
                                varWeekStart = varWeekEnd.AddDays(1);

                                //Get the end of the period
                                if (varPeriod == "week")
                                    varWeekEnd = varWeekStart.AddDays(6);
                                else if (varPeriod == "month")
                                {
                                    varWeekEnd = varWeekStart.AddMonths(1);
                                    varWeekEnd = varWeekEnd.AddDays(-1);
                                }
                                else
                                    varWeekEnd = varYearEnd;

                                //Add to the next period number
                                varWeek++;
                            } //while loop: weekstart < year end

                            // Create the report part of the extract
                            sbHeaderList.AppendLine("<p>");
                            sbHeaderList.AppendLine("<i>Note: Minor rounding issues with future remaining amounts</i><br>");
                            sbHeaderList.AppendLine("<table border=0 cellpadding=3>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td>&nbsp;</td>");
                            sbHeaderList.AppendLine("<td>&nbsp;</td>");
                            sbHeaderList.AppendLine("<td>&nbsp;</td>");
                            sbHeaderList.AppendLine("<td>&nbsp;</td>");
                            sbHeaderList.AppendLine("<td>&nbsp;</td>");
                            sbHeaderList.AppendLine("<td>Episodes</td>");
                            sbHeaderList.AppendLine("<td>" + varPeriod + " Spent</td>");
                            sbHeaderList.AppendLine("<td>" + varPeriod + " Remaining</td>");
                            sbHeaderList.AppendLine("<td>Total Spent</td>");
                            sbHeaderList.AppendLine("<td>Total Remaining</td>");
                            sbHeaderList.AppendLine("<td>&nbsp;</td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine(sbFinancialList.ToString());
                            sbHeaderList.AppendLine("</table>");

                            sbReport2List.AppendLine("</table>");

                        }
                        // Do this if the report type is a spend report
                        else if (varReport == "Spend Report")
                        {
                            // Get the financial report name and file name
                            varReportName = varBrokerage + " spend report for " + varMonthStart.ToString("MMM") + " " + varMonthStart.Year;
                            varFileName = varBrokerage + " spend report for " + varMonthStart.ToString("MMM") + " " + varMonthStart.Year + ".xls";

                            varFileName2 = varFileName.Replace(".xls", ".txt");
                            varFileName2 = "Errors for " + varFileName2;

                            // Fetch statements for database
                            // Financial data for period and brokerage chosen
                            dbFinancialList = @"
                            <fetch version='1.0' mapping='logical' distinct='false' >
                                <entity name='new_palmclientfinancial' >
                                    <attribute name='new_palmclientfinancialid' />
                                    <attribute name='new_entrydate' />
                                    <attribute name='new_datepaid' />
                                    <attribute name='new_cheque' />
                                    <attribute name='new_voucherid' />
                                    <attribute name='new_purpose' />
                                    <attribute name='new_payee' />
                                    <attribute name='new_amount' />
                                    <attribute name='new_gst' />
                                    <attribute name='new_paid' />
                                    <attribute name='new_genfinid' />
                                    <attribute name='new_paidby' />
                                    <attribute name='new_firstnamenc' />
                                    <attribute name='new_surnamenc' />
                                    <attribute name='new_firstname' />
                                    <attribute name='new_surname' />
                                    <filter type='and' >
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + varMonthStart + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='le' value='" + varMonthEnd + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_paidby' operator='ne' value='100000003' />
                                        
                                    </filter>
                                    <link-entity name='new_palmsuassistance' from='new_palmsuassistanceid' to='new_assistance' link-type='outer' >
                                        <attribute name='new_code' />
                                    </link-entity>
                                    <link-entity name='new_palmsubrokerage' from='new_palmsubrokerageid' to='new_brokerage' link-type='outer' >
                                        <attribute name = 'new_bankacc' />
                                     </link-entity>
                                </entity>
                            </fetch> ";
                            //fin, 1:assistance, 2:brokerage


                            // Total spend for period and brokerage chosen
                            dbSpendList = @"
                                <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                    <entity name='new_palmclientfinancial'>
                                    <attribute name='new_amount' alias='totalamount_sum' aggregate='sum'/> 
                                    <filter type='and'>
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + varYearStart + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='lt' value='" + varMonthStart + @"' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_paidby' operator='ne' value='100000003' />
                                    </filter>
                                    </entity>
                                </fetch> ";

                            // Budget for period and brokerage chosen
                            dbBudgetList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                                <entity name='new_palmsubudget'>
                                <attribute name='new_amount'/>
                                <attribute name='new_startdate'/>
                                <attribute name='new_enddate'/>
                                <attribute name='new_palmsubudgetid'/>
                                <filter type='and'>
                                    <condition entityname='new_palmsubudget' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                    <condition entityname='new_palmsubudget' attribute='new_startdate' operator='le' value='" + varYearStart + @"' />
                                    <condition entityname='new_palmsubudget' attribute='new_enddate' operator='ge' value='" + varYearStart + @"' />
                                </filter>
                                </entity>
                            </fetch> ";

                            // Get the fetch XML data and place in entity collection objects
                            result = _service.RetrieveMultiple(new FetchExpression(dbFinancialList));
                            result2 = _service.RetrieveMultiple(new FetchExpression(dbSpendList));
                            result3 = _service.RetrieveMultiple(new FetchExpression(dbBudgetList));

                            // Loop through financial data
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
                                if (c.FormattedValues.Contains("new_palmclientfinancialid"))
                                    dbPalmClientFinancialId = c.FormattedValues["new_palmclientfinancialid"];
                                else if (c.Attributes.Contains("new_palmclientfinancialid"))
                                    dbPalmClientFinancialId = c.Attributes["new_palmclientfinancialid"].ToString();
                                else
                                    dbPalmClientFinancialId = "";

                                if (c.FormattedValues.Contains("new_entrydate"))
                                    dbEntryDate = c.FormattedValues["new_entrydate"];
                                else if (c.Attributes.Contains("new_entrydate"))
                                    dbEntryDate = c.Attributes["new_entrydate"].ToString();
                                else
                                    dbEntryDate = "";

                                // Convert date from American format to Australian format
                                dbEntryDate = cleanDateAM(dbEntryDate);

                                if (c.FormattedValues.Contains("new_cheque"))
                                    dbCheque = c.FormattedValues["new_cheque"];
                                else if (c.Attributes.Contains("new_cheque"))
                                    dbCheque = c.Attributes["new_cheque"].ToString();
                                else
                                    dbCheque = "";


                                if (c.FormattedValues.Contains("new_voucherid"))
                                    dbVoucher = c.FormattedValues["new_voucherid"];
                                else if (c.Attributes.Contains("new_voucherid"))
                                    dbVoucher = c.Attributes["new_voucherid"].ToString();
                                else
                                    dbVoucher = "";

                                if (c.FormattedValues.Contains("new_purpose"))
                                    dbPurpose = c.FormattedValues["new_purpose"];
                                else if (c.Attributes.Contains("new_purpose"))
                                    dbPurpose = c.Attributes["new_purpose"].ToString();
                                else
                                    dbPurpose = "";

                                if (c.FormattedValues.Contains("new_payee"))
                                    dbPayee = c.FormattedValues["new_payee"];
                                else if (c.Attributes.Contains("new_payee"))
                                    dbPayee = c.Attributes["new_payee"].ToString();
                                else
                                    dbPayee = "";

                                if (c.FormattedValues.Contains("new_amount"))
                                    Double.TryParse(c.FormattedValues["new_amount"], out dbAmount);
                                else if (c.Attributes.Contains("new_amount"))
                                    Double.TryParse(c.Attributes["new_amount"].ToString(), out dbAmount);
                                else
                                    dbAmount = 0;

                                if (c.FormattedValues.Contains("new_gst"))
                                    Double.TryParse(c.FormattedValues["new_gst"], out dbGst);
                                else if (c.Attributes.Contains("new_gst"))
                                    Double.TryParse(c.Attributes["new_gst"].ToString(), out dbGst);
                                else
                                    dbGst = 0;

                                if (c.FormattedValues.Contains("new_firstnamenc"))
                                    dbFirstNameNc = c.FormattedValues["new_firstnamenc"];
                                else if (c.Attributes.Contains("new_firstnamenc"))
                                    dbFirstNameNc = c.Attributes["new_firstnamenc"].ToString();
                                else
                                    dbFirstNameNc = "";

                                if (c.FormattedValues.Contains("new_firstname"))
                                    dbFirstName = c.FormattedValues["new_firstname"];
                                else if (c.Attributes.Contains("new_firstname"))
                                    dbFirstName = c.Attributes["new_firstname"].ToString();
                                else
                                    dbFirstName = "";

                                if (String.IsNullOrEmpty(dbFirstName) == true)
                                    dbFirstName = dbFirstNameNc;

                                if (c.FormattedValues.Contains("new_surnamenc"))
                                    dbSurnameNc = c.FormattedValues["new_surnamenc"];
                                else if (c.Attributes.Contains("new_surnamenc"))
                                    dbSurnameNc = c.Attributes["new_surnamenc"].ToString();
                                else
                                    dbSurnameNc = "";

                                if (c.FormattedValues.Contains("new_surname"))
                                    dbSurname = c.FormattedValues["new_surname"];
                                else if (c.Attributes.Contains("new_surname"))
                                    dbSurname = c.Attributes["new_surname"].ToString();
                                else
                                    dbSurname = "";

                                if (String.IsNullOrEmpty(dbSurname) == true)
                                    dbSurname = dbSurnameNc;



                                if (c.Attributes.Contains("new_palmsuassistance1.new_code"))
                                    dbAssistanceCode = c.GetAttributeValue<AliasedValue>("new_palmsuassistance1.new_code").Value.ToString();
                                else
                                    dbAssistanceCode = "";

                                if (c.Attributes.Contains("new_palmsubrokerage2.new_bankacc"))
                                    dbAccountNum = c.GetAttributeValue<AliasedValue>("new_palmsubrokerage2.new_bankacc").Value.ToString();
                                else
                                    dbAccountNum = "";


                                if (c.FormattedValues.Contains("new_paidby"))
                                    dbPaidBy = c.FormattedValues["new_paidby"];
                                else if (c.Attributes.Contains("new_paidby"))
                                    dbPaidBy = c.Attributes["new_paidby"].ToString();
                                else
                                    dbPaidBy = "";

                                if (c.FormattedValues.Contains("new_genfinid"))
                                    dbGeneratedId = c.FormattedValues["new_genfinid"];
                                else if (c.Attributes.Contains("new_genfinid"))
                                    dbGeneratedId = c.Attributes["new_genfinid"].ToString();
                                else
                                    dbGeneratedId = "";

                                // Append line to report
                                if (dbPaidBy == "Cheque")
                                    sbReportList.AppendLine("<tr>\r\n<td class=\"prBorder\">" + dbEntryDate + "</td>\r\n<td class=\"prBorder\">" + dbPaidBy + "</td>\r\n<td class=\"prBorder\">" + dbCheque + "</td>\r\n<td class=\"prBorder\">" + dbFirstName + " " + dbSurname + "</td>\r\n<td class=\"prBorder\">" + dbPurpose + "</td>\r\n<td class=\"prBorder\">" + dbPayee + "</td>\r\n<td class=\"prBorder\">" + dbAmount.ToString("C") + "</td>\r\n<td class=\"prBorder\">" + dbGst.ToString("C") + "</td>\r\n<td class=\"prBorder\"> 1-" + dbAssistanceCode + "</td>\r\n<td class=\"prBorder\">" + dbAccountNum + "</td>\r\n</tr>");
                                else if (dbPaidBy == "Voucher")
                                    sbReportList.AppendLine("<tr>\r\n<td class=\"prBorder\">" + dbEntryDate + "</td>\r\n<td class=\"prBorder\">" + dbPaidBy + "</td>\r\n<td class=\"prBorder\">" + dbVoucher + "</td>\r\n<td class=\"prBorder\">" + dbFirstName + " " + dbSurname + "</td>\r\n<td class=\"prBorder\">" + dbPurpose + "</td>\r\n<td class=\"prBorder\">" + dbPayee + "</td>\r\n<td class=\"prBorder\">" + dbAmount.ToString("C") + "</td>\r\n<td class=\"prBorder\">" + dbGst.ToString("C") + "</td>\r\n<td class=\"prBorder\"> 1-" + dbAssistanceCode + "</td>\r\n<td class=\"prBorder\">" + dbAccountNum + "</td>\r\n</tr>");
                                else
                                    sbReportList.AppendLine("<tr>\r\n<td class=\"prBorder\">" + dbEntryDate + "</td>\r\n<td class=\"prBorder\">" + dbPaidBy + "</td>\r\n<td class=\"prBorder\">" + dbGeneratedId + "</td>\r\n<td class=\"prBorder\">" + dbFirstName + " " + dbSurname + "</td>\r\n<td class=\"prBorder\">" + dbPurpose + "</td>\r\n<td class=\"prBorder\">" + dbPayee + "</td>\r\n<td class=\"prBorder\">" + dbAmount.ToString("C") + "</td>\r\n<td class=\"prBorder\">" + dbGst.ToString("C") + "</td>\r\n<td class=\"prBorder\"> 1-" + dbAssistanceCode + "</td>\r\n<td class=\"prBorder\">" + dbAccountNum + "</td>\r\n</tr>");

                                //sbReportList.AppendLine("<tr>\r\n<td class=\"prBorder\">" + dbEntryDate + "</td>\r\n<td class=\"prBorder\">" + dbCheque + "</td>\r\n<td class=\"prBorder\">" + dbFirstName + " " + dbSurname + "</td>\r\n<td class=\"prBorder\">" + dbPurpose + "</td>\r\n<td class=\"prBorder\">" + dbPayee + "</td>\r\n<td class=\"prBorder\">" + dbAmount.ToString("C") + "</td>\r\n<td class=\"prBorder\">" + dbGst.ToString("C") + "</td>\r\n<td class=\"prBorder\"> 1-" + dbAssistanceCode + "</td>\r\n</tr>");

                                //Add amount to the total
                                varTotal += dbAmount;
                            } // financial loop

                            // Get total spend for period
                            foreach (var c in result2.Entities)
                            {
                                if (c.FormattedValues.Contains("totalamount_sum"))
                                    Double.TryParse(c.FormattedValues["totalamount_sum"], out dbYtdSpend);
                                else if (c.Attributes.Contains("totalamount_sum"))
                                    Double.TryParse(c.Attributes["totalamount_sum"].ToString(), out dbYtdSpend);
                                else
                                    dbYtdSpend = 0;
                            } // spend loop

                            // Get budget for period
                            foreach (var c in result3.Entities)
                            {
                                if (c.FormattedValues.Contains("new_amount"))
                                    Double.TryParse(c.FormattedValues["new_amount"], out varAmount);
                                else if (c.Attributes.Contains("new_amount"))
                                    Double.TryParse(c.Attributes["new_amount"].ToString(), out varAmount);
                                else
                                    varAmount = 0;
                            } // spend loop


                            //Get the YTD budget as at end of the month
                            ts1 = varYearEnd - varYearStart;
                            ts2 = varMonthEnd - varYearStart;
                            varYTDB = (double)varAmount / (ts1.TotalDays + 1) * (ts2.TotalDays + 1);


                            //Header part of the Report extract
                            sbHeaderList.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                            sbHeaderList.AppendLine("<head>");
                            sbHeaderList.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                            sbHeaderList.AppendLine("<!--[if gte mso 9]>");
                            sbHeaderList.AppendLine("<xml>");
                            sbHeaderList.AppendLine("<x:ExcelWorkbook>");
                            sbHeaderList.AppendLine("<x:ExcelWorksheets>");
                            sbHeaderList.AppendLine("<x:ExcelWorksheet>");

                            //this line names the worksheet
                            sbHeaderList.AppendLine("<x:Name>Financial Data</x:Name>");

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

                            // Create spend report
                            sbHeaderList.AppendLine("<table width=\"100%\" border=0 cellpadding=5 class=\"myClass1\">");
                            sbHeaderList.AppendLine("<tr><td colspan=10 align=\"center\" class=\"fontTitle\">" + varBrokerage + "</td></tr>");
                            sbHeaderList.AppendLine("<tr><td colspan=10> </td></tr>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td colspan=10 align=\"center\" class=\"fontNormal\">BUDGET FOR " + varYearStart.Year + "/" + varYearEnd.Year + " | " + varAmount.ToString("C") + "</td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine("<tr><td colspan=10> </td></tr>");
                            sbHeaderList.AppendLine("<tr><td colspan=10 align=\"center\" class=\"fontNormal\">Payments Report for the month of " + varMonthStart.ToString("MMMM") + " " + varMonthStart.Year + "</td></tr>");
                            sbHeaderList.AppendLine("<tr><td colspan=10> </td></tr>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td colspan=2> </td>");
                            sbHeaderList.AppendLine("<td colspan=2 class=\"fontNormal\">Budget available</td>");
                            sbHeaderList.AppendLine("<td class=\"fontNormal\">" + (varAmount - dbYtdSpend).ToString("C") + "</td>");
                            sbHeaderList.AppendLine("<td colspan=2> </td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine("<tr><td colspan=10> </td></tr>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td class=\"prBorderNW\">Date</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorderNW\">Payment Method</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorderNW\">Payment ID Number</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\">Client Name</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorderNW\">Purpose</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\">Payee</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorderNW\">Amount</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorderNW\">GST</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorderNW\">Assistance Code</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorderNW\">Account ID</td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine(sbReportList.ToString());
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\">Total</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\"> </td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\"> </td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\"> </td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\"> </td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\"> </td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\">" + varTotal.ToString("C") + "</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\"> </td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine("<tr><td colspan=10> </td></tr>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td colspan=2> </td>");
                            sbHeaderList.AppendLine("<td colspan=2 class=\"fontNormal\">Budget Remaining</td>");
                            sbHeaderList.AppendLine("<td class=\"fontNormal\">" + (varAmount - dbYtdSpend - varTotal).ToString("C") + "</td>");
                            sbHeaderList.AppendLine("<td colspan=2> </td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine("<tr><td colspan=10> </td></tr>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td colspan=2 class=\"fontNormal\">Year to Date Actual</td>");
                            sbHeaderList.AppendLine("<td colspan=5 class=\"fontNormal\">" + (dbYtdSpend + varTotal).ToString("C") + "</td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td colspan=2 class=\"fontNormal\">Year to Date Budget</td>");
                            sbHeaderList.AppendLine("<td colspan=5 class=\"fontNormal\">" + (varYTDB).ToString("C") + "</td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine("<tr><td colspan=10> </td></tr>");
                            sbHeaderList.AppendLine("<tr>");
                            sbHeaderList.AppendLine("<td colspan=2 class=\"fontNormal\">Variance</td>");
                            sbHeaderList.AppendLine("<td colspan=5 class=\"fontNormal\">" + (varYTDB - (dbYtdSpend + varTotal)).ToString("C") + "</td>");
                            sbHeaderList.AppendLine("</tr>");
                            sbHeaderList.AppendLine("</table>");

                        }

                        //DHHS REPORT SIX MONTH BASIS HEF & PRAP

                        // Do this if the report type is non ship user report
                        else if (varReport == "Non SHIP User Report" || varReport == "Non SHIP User Report - GST and Fund")
                        {
                            string[] agencyArray = { "Haven Home Safe - H2H Loddon", "Haven - Kyabram Saap", "Haven - Emergency Accom Options for men", "Haven - Supporting Families at Risk Mallee", "Haven - PRAP Bdgo", "Haven - PRAP Mlda", "Haven Home Safe - Prison IAP", "Haven Home Safe - Tenancy Plus Program", "Haven Home Safe - H2H Barwon", "RSAP Assertive Outreach Bendigo", "RSAP Assertive Outreach Swan Hill", "Haven - Bendigo Saap", "Haven - Youth support services", "Haven - THM Assessment and Planning", "Haven Home Safe - Metro IAP" };
                            for (int i = 0; i < agencyArray.Length; i++)
                            {
                                // Get the report name and file name
                                //varReportName = varBrokerage + " extract for " + cleanDate(varStartDate) + " to " + cleanDate(varEndDate);
                                //varFileName = varBrokerage + " extract for " + cleanDate(varStartDate) + " to " + cleanDate(varEndDate) + ".xls";

                                sbHeaderList.Clear();
                                sbReportList.Clear();

                                //Clear variables.
                                varFileName = "No Report"; // Report file name
                                varReportName = "No Report"; // Name of report
                                dbPalmClientFinancialId = "";
                                dbEntryDate = "";
                                dbPaidDate = "";
                                dbPurpose = "";
                                dbPayee = "";
                                dbAmount = 0;
                                //varAmount = 0;
                                dbGst = 0;
                                dbPurchaseType = "";
                                varNMDStype = "";
                                dbPalmClientId = "";
                                dbClient = "";
                                dbDob = "";
                                dbDobEst = "";
                                dbDobEstD = ""; // DOB day estimated field
                                dbDobEstM = ""; // DOB month estimated field
                                dbDobEstY = ""; // DOB year estimated field
                                dbShorSlk = "";
                                dbWeekRent = "";
                                dbPalmClientSupportId = "";
                                dbPalmClientSupportIdOld = ""; //Support Period id - old
                                dbDoShor = "";
                                dbShorAgency = "";
                                dbLocality = "";
                                dbState = "";
                                dbPostcode = "";
                                dbPuhId = "";
                                dbBrokerage = "";
                                dbPuhSupportPeriod = "";
                                dbPuhIdOld = ""; //Presenting unit head old
                                dbPuhClient = "";
                                //Variable for setting old ids for imported records.
                                varPrintSpId = "";
                                varPrintPuhId = "";
                                // Variables for SLK
                                varSLK = "";
                                varSurname = "";
                                varFirstName = "";
                                varDob = "";
                                varGender = "";
                                varDobFlag = ""; // Dob flag


                                // Fetch statements for database
                                // Get the financial data for the period and brokerage chosen
                                dbFinancialList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclient'>
                                <attribute name='new_palmclientid' />
                                <attribute name='new_address' />
                                <attribute name='new_firstname' />
                                <attribute name='new_surname' />
                                <attribute name='new_gender' />
                                <attribute name='new_dob' />
                                <attribute name='new_dobest' />
                                <attribute name='new_dobestd' />
                                <attribute name='new_dobestm' />
                                <attribute name='new_dobesty' />
                                <attribute name='new_mdsslk' />
                                <attribute name='new_shorslk' />
                                <attribute name='new_weekrent' />
                                <link-entity name='new_palmclientsupport' to='new_palmclientid' from='new_client' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_supportperiodidold' />
                                    <attribute name='new_doshor' />
                                    <attribute name='new_locality' />
                                    <attribute name='new_puhid' />
                                    <attribute name='new_puhidold' />
                                    <link-entity name='new_palmclientfinancial' to='new_palmclientsupportid' from='new_supportperiod' link-type='inner'>
                                        <attribute name='new_entrydate' />
                                        <attribute name='new_datepaid' />
                                        <attribute name='new_amount' />
                                        <attribute name='new_gst' />
                                        <attribute name='new_assistance' />
                                        <attribute name='new_purhchasetype' />
                                        <attribute name='new_brokerage' />
                                        <attribute name='new_payee' />
                                        <link-entity name='new_palmsubrokerage' to='new_brokerage' from='new_palmsubrokerageid' link-type='outer'>
                                            <attribute name='new_shorthand' />
                                        </link-entity>
                                    </link-entity>
                                    <link-entity name='new_palmddllocality' to='new_locality' from='new_palmddllocalityid' link-type='outer'>
                                        <attribute name='new_postcode' />
                                        <attribute name='new_state' />
                                    </link-entity>
                                    <link-entity name='new_palmsushoragency' to='new_doshor' from='new_palmsushoragencyid' link-type='inner'>
                                        <attribute name='new_agency' />
                                    </link-entity>
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientfinancial' attribute='new_datepaid' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_datepaid' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmsushoragency' attribute='new_agency' operator='eq' value='" + agencyArray[i] + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_amount' operator='gt' value='1' />
                                    <filter type='or'>
                                        <condition entityname='new_palmsubrokerage' attribute='new_shorthand' operator='like' value='hef%' />
                                        <condition entityname='new_palmsubrokerage' attribute='new_shorthand' operator='like' value='prap%' />
                                    </filter>
                                </filter>
                              </entity>
                            </fetch> ";


                                //<condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                /*
                                                                     <filter type='or'>
                                        <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='like' value='hef%' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='like' value='prap%' />
                                    </filter>
                                 */

                                // Get the fetch XML data and place in entity collection objects
                                result = _service.RetrieveMultiple(new FetchExpression(dbFinancialList));

                                // Loop through financial data
                                foreach (var c in result.Entities)
                                {
                                    varGSTUpDown = "";
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

                                    if (String.IsNullOrEmpty(dbFirstName) == true)
                                        dbFirstName = "999999";

                                    if (c.FormattedValues.Contains("new_surname"))
                                        dbSurname = c.FormattedValues["new_surname"];
                                    else if (c.Attributes.Contains("new_surname"))
                                        dbSurname = c.Attributes["new_surname"].ToString();
                                    else
                                        dbSurname = "";

                                    if (String.IsNullOrEmpty(dbSurname) == true)
                                        dbSurname = "999999";

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

                                    if (String.IsNullOrEmpty(dbDob) == true)
                                        dbDob = "1-Jan-1970";

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
                                            varDobFlag = dbDobEstD.Substring(0, 1) + varDobFlag.Substring(1, 2);
                                        if (string.IsNullOrEmpty(dbDobEstM) == false)
                                            varDobFlag = varDobFlag.Substring(0, 1) + dbDobEstM.Substring(0, 1) + varDobFlag.Substring(2, 1);
                                        if (string.IsNullOrEmpty(dbDobEstY) == false)
                                            varDobFlag = varDobFlag.Substring(0, 2) + dbDobEstY.Substring(0, 1);
                                    }
                                    else if (dbDobEst == "Yes")
                                        // Old format - If estimated is yes then do estimated for all
                                        varDobFlag = "EEE";

                                    // Get new dog flag format
                                    dbDobEst = varDobFlag;


                                    if (c.FormattedValues.Contains("new_mdsslk"))
                                        dbMdsSlk = c.FormattedValues["new_mdsslk"];
                                    else if (c.Attributes.Contains("new_mdsslk"))
                                        dbMdsSlk = c.Attributes["new_mdsslk"].ToString();
                                    else
                                        dbMdsSlk = "";

                                    if (c.FormattedValues.Contains("new_shorslk"))
                                        dbShorSlk = c.FormattedValues["new_shorslk"];
                                    else if (c.Attributes.Contains("new_shorslk"))
                                        dbShorSlk = c.Attributes["new_shorslk"].ToString();
                                    else
                                        dbShorSlk = "";

                                    if (c.FormattedValues.Contains("new_weekrent"))
                                        dbWeekRent = c.FormattedValues["new_weekrent"];
                                    else if (c.Attributes.Contains("new_weekrent"))
                                        dbWeekRent = c.Attributes["new_weekrent"].ToString();
                                    else
                                        dbWeekRent = "";

                                    dbWeekRent = cleanString(dbWeekRent, "double");

                                    Double.TryParse(dbWeekRent, out varCheckDouble);
                                    dbWeekRent = varCheckDouble.ToString("C");

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
                                        if (dbPalmClientSupportIdOld.Length < 5) //Support period id's from SHIP are 7 characters - doubtful a palm client would have had more than 10000 support periods so roughly works.
                                            varPrintSpId = dbClient + "_" + dbPalmClientSupportIdOld;
                                        else
                                            varPrintSpId = dbPalmClientSupportIdOld;
                                    }
                                    else
                                        varPrintSpId = dbPalmClientSupportId;


                                    if (c.FormattedValues.Contains("new_palmclientsupport1.new_doshor"))
                                        dbDoShor = c.FormattedValues["new_palmclientsupport1.new_doshor"];
                                    else if (c.Attributes.Contains("new_palmclientsupport1.new_doshor"))
                                        dbDoShor = c.Attributes["new_palmclientsupport1.new_doshor"].ToString();
                                    else
                                        dbDoShor = "";

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

                                    if (c.FormattedValues.Contains("new_palmclientsupport1.new_puhid"))
                                        dbPuhId = c.FormattedValues["new_palmclientsupport1.new_puhid"];
                                    else if (c.Attributes.Contains("new_palmclientsupport1.new_puhid"))
                                        dbPuhId = c.GetAttributeValue<AliasedValue>("new_palmclientsupport1.new_puhid").Value.ToString();
                                    else
                                        dbPuhId = dbPalmClientSupportId;

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

                                    if (c.FormattedValues.Contains("new_palmclientfinancial2.new_entrydate"))
                                        dbEntryDate = c.FormattedValues["new_palmclientfinancial2.new_entrydate"];
                                    else if (c.Attributes.Contains("new_palmclientfinancial2.new_entrydate"))
                                        dbEntryDate = c.Attributes["new_palmclientfinancial2.new_entrydate"].ToString();
                                    else
                                        dbEntryDate = "";

                                    if (c.FormattedValues.Contains("new_palmclientfinancial2.new_datepaid"))
                                        dbPaidDate = c.FormattedValues["new_palmclientfinancial2.new_datepaid"];
                                    else if (c.Attributes.Contains("new_palmclientfinancial2.new_datepaid"))
                                        dbPaidDate = c.Attributes["new_palmclientfinancial2.new_datepaid"].ToString();
                                    else
                                        dbPaidDate = "";

                                    // Convert date from American format to Australian format
                                    dbEntryDate = cleanDateAM(dbEntryDate);
                                    dbPaidDate = cleanDateAM(dbPaidDate);

                                    if (c.FormattedValues.Contains("new_palmclientfinancial2.new_amount"))
                                        Double.TryParse(c.FormattedValues["new_palmclientfinancial2.new_amount"], out dbAmount);
                                    else if (c.Attributes.Contains("new_palmclientfinancial2.new_amount"))
                                        Double.TryParse(c.Attributes["new_palmclientfinancial2.new_amount"].ToString(), out dbAmount);
                                    else
                                        dbAmount = 0;

                                    if (c.FormattedValues.Contains("new_palmclientfinancial2.new_gst"))
                                        Double.TryParse(c.FormattedValues["new_palmclientfinancial2.new_gst"], out dbGst);
                                    else if (c.Attributes.Contains("new_palmclientfinancial2.new_gst"))
                                        Double.TryParse(c.Attributes["new_palmclientfinancial2.new_gst"].ToString(), out dbGst);
                                    else
                                        dbGst = 0;


                                    //WORK OUT IF GST IS 10% of amount or not.
                                    if (dbAmount * 0.1 == dbGst)
                                        varGSTUpDown = "Amount is inclusive of GST";
                                    else if (dbGst == 0)
                                        varGSTUpDown = "Unknown";
                                    else
                                        varGSTUpDown = "Amount is exclusive of GST";

                                    if (c.FormattedValues.Contains("new_palmclientfinancial2.new_assistance"))
                                        dbAssistance = c.FormattedValues["new_palmclientfinancial2.new_assistance"];
                                    else if (c.Attributes.Contains("new_palmclientfinancial2.new_assistance"))
                                        dbAssistance = c.Attributes["new_palmclientfinancial2.new_assistance"].ToString();
                                    else
                                        dbAssistance = "";

                                    if (c.FormattedValues.Contains("new_palmclientfinancial2.new_brokerage"))
                                        dbBrokerage = c.FormattedValues["new_palmclientfinancial2.new_brokerage"];
                                    else if (c.Attributes.Contains("new_palmclientfinancial2.new_brokerage"))
                                        dbBrokerage = c.Attributes["new_palmclientfinancial2.new_brokerage"].ToString();
                                    else
                                        dbBrokerage = "";

                                    if (varReport == "Non SHIP User Report")
                                    {
                                        if (String.IsNullOrEmpty(dbBrokerage) == false)
                                            if (dbBrokerage.ToLower().IndexOf("prap") >= 0)
                                                dbBrokerage = "Private Rental Assistance Program";
                                            else if (dbBrokerage.ToLower().IndexOf("hef") >= 0)
                                                dbBrokerage = "Housing Establishment Fund";
                                            else
                                                dbBrokerage = "";
                                    }

                                    if (c.FormattedValues.Contains("new_palmclientfinancial2.new_purhchasetype"))
                                        dbPurchaseType = c.FormattedValues["new_palmclientfinancial2.new_purhchasetype"];
                                    else if (c.Attributes.Contains("new_palmclientfinancial2.new_purhchasetype"))
                                        dbPurchaseType = c.Attributes["new_palmclientfinancial2.new_purhchasetype"].ToString();
                                    else
                                        dbPurchaseType = "";

                                    string[] purchaseTypesShortTerm = { "motels/hotels", "private rooming house", "caravan park", "other" };
                                    string[] purchaseTypesEstablishing = { "rent to establish a tenancy (ria)", "rent to establish a tenancy", "bond", "bond loan debt", "landlord incentive", "removalists or storage", "household items to establish a tenancy", "domestic safety measures to establish a tenancy", "property modifications to establish a tenancy", "other establishing a tenancy" };
                                    string[] purchaseTypesMaintaining = { "rent to maintain a tenancy (arr)", "rent to maintain a tenancy", "clean up", "household items to maintain a tenancy", "whitegoods to maintain a tenancy", "property modifications to maintain a tenancy", "domestic safety measures to maintain a tenancy", "other maintaining a tenancy" };
                                    string[] purchaseTypesEducation = { "primary and high school costs (fees, resources, uniforms, laptop etc)", "Other training/education/employment (fees, travel, interviews etc)" };
                                    string[] purchaseTypesSpecialist = { "cultural safety and strenghtening", "child care (includes fees, payments to informal carer etc)", "legal services", "professional counselling/psychological/parenting support, etc", "medical and health" };
                                    string[] purchaseTypesOther = { "food purchases", "food vouchers", "other safety measures", "medical/pharmaceuticals/aids", "other health and wellbeing supports (sport, support groups, holidays, respite, cultural etc)", "connectivity/phones/computers", "utilities", "clothes", "travel/commuting (includes cars/driving lessons, myki cards etc)", "other payment" };

                                    if (Array.IndexOf(purchaseTypesShortTerm, dbPurchaseType.ToLower()) >= 0)
                                        varNMDStype = "Payment for short term or emergency accommodation";
                                    else if (Array.IndexOf(purchaseTypesEstablishing, dbPurchaseType.ToLower()) >= 0)
                                        varNMDStype = "Establishing a tenancy";
                                    else if (Array.IndexOf(purchaseTypesMaintaining, dbPurchaseType.ToLower()) >= 0)
                                        varNMDStype = "Maintaining a tenancy";
                                    else if (Array.IndexOf(purchaseTypesEducation, dbPurchaseType.ToLower()) >= 0)
                                        varNMDStype = "Payment for training/education/employment";
                                    else if (Array.IndexOf(purchaseTypesSpecialist, dbPurchaseType.ToLower()) >= 0)
                                        varNMDStype = "Payment for accessing external specialist services";
                                    else if (Array.IndexOf(purchaseTypesEstablishing, dbPurchaseType.ToLower()) >= 0)
                                        varNMDStype = "Other Payments";
                                    else
                                        varNMDStype = "";

                                    if (c.FormattedValues.Contains("new_palmclientfinancial2.new_payee"))
                                        dbPayee = c.FormattedValues["new_palmclientfinancial2.new_payee"];
                                    else if (c.Attributes.Contains("new_palmclientfinancial2.new_payee"))
                                        dbPayee = c.GetAttributeValue<AliasedValue>("new_palmclientfinancial2.new_payee").Value.ToString();
                                    else
                                        dbPayee = "";

                                    if (c.FormattedValues.Contains("new_palmsushoragency5.new_agency"))
                                        dbShorAgency = c.FormattedValues["new_palmsushoragency5.new_agency"];
                                    else if (c.Attributes.Contains("new_palmsushoragency5.new_agency"))
                                        dbShorAgency = c.GetAttributeValue<AliasedValue>("new_palmsushoragency5.new_agency").Value.ToString();
                                    else
                                        dbShorAgency = "";

                                    //Sorts out issue with data from palm where Payee contains paid by information.
                                    int payeeIndex = dbPayee.ToLower().IndexOf("paid");
                                    if (payeeIndex > 0)
                                        dbPayee = dbPayee.Substring(0, payeeIndex - 1);

                                    // Create the SLK based on firstname, surname, gender and dob
                                    varSurname = dbSurname.ToUpper() + "22222";
                                    varSurname = varSurname.Substring(1, 1) + varSurname.Substring(2, 1) + varSurname.Substring(4, 1);

                                    if (varSurname == "222")
                                        varSurname = "999";

                                    varFirstName = dbFirstName.ToUpper() + "22222";
                                    varFirstName = varFirstName.Substring(1, 1) + varFirstName.Substring(2, 1);

                                    if (varFirstName == "22")
                                        varFirstName = "99";

                                    //Get the gender code
                                    if (dbGender == "Female")
                                        varGender = "2";
                                    else
                                        varGender = "1";

                                    //Put dob into expected format
                                    varDob = cleanDateS(Convert.ToDateTime(dbDob));

                                    //Get the statistical linkage key
                                    varSLK = varSurname + varFirstName + varDob + varGender;

                                    // Append data to report
                                    sbReportList.Append("<tr>\r\n<td>" + dbShorAgency + "</td>\r\n<td>" + dbDoShor + "</td>\r\n<td>" + varPrintPuhId + "</td>\r\n<td>" + varPrintSpId + "</td>\r\n<td>" + varDobFlag + "</td>\r\n<td>" + varSLK + "</td>\r\n<td>" + dbShorSlk + "</td>\r\n<td>" + dbPaidDate + "</td>\r\n<td>" + dbPayee + "</td>\r\n<td>" + dbFirstName + "</td>\r\n<td>" + dbSurname + "</td>\r\n<td>" + dbAmount + "</td><td>");

                                    if (varReport == "Non SHIP User Report - GST and Fund")
                                    {
                                        sbReportList.Append(dbGst + "</td>\r\n <td>" + varGSTUpDown + "</td>\r\n<td>");
                                    }

                                    sbReportList.Append(dbBrokerage);

                                    sbReportList.AppendLine("</td>\r\n<td>&nbsp;</td>\r\n<td>&nbsp;</td>\r\n<td>&nbsp;</td>\r\n<td>" + dbPurchaseType + "</td>\r\n<td>" + varNMDStype + "</td>\r\n<td>&nbsp;</td>\r\n<td>" + dbLocality + "</td>\r\n<td>" + dbWeekRent + "</td>\r\n<td>&nbsp;</td>\r\n</tr>");

                                } // client loop        

                                varReportName = dbDoShor + " " + agencyArray[i] + " " + cleanDate(varStartDate);
                                varFileName = dbDoShor + " " + agencyArray[i] + " " + cleanDate(varStartDate) + ".xls";

                                //Header part of the Report
                                sbHeaderList.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                                sbHeaderList.AppendLine("<head>");
                                sbHeaderList.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                                sbHeaderList.AppendLine("<!--[if gte mso 9]>");
                                sbHeaderList.AppendLine("<xml>");
                                sbHeaderList.AppendLine("<x:ExcelWorkbook>");
                                sbHeaderList.AppendLine("<x:ExcelWorksheets>");
                                sbHeaderList.AppendLine("<x:ExcelWorksheet>");

                                //this line names the worksheet
                                sbHeaderList.AppendLine("<x:Name>Financial Data</x:Name>");

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
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">SHS Name</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">SHS Agency ID</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Household (PUH) ID</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Support period ID</td>"); // LOOK AT
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">DOB Accuracy</td>"); // NEW
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">SLK</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Alphacode</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Date of payment (Date Paid)</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Payee</td>"); // NEW
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">First Name</td>"); // NEW
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Last Name</td>"); // NEW
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Amount $</td>");
                                if (varReport == "Non SHIP User Report - GST and Fund")
                                {
                                    sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">GST $</td>");
                                    sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">GST Inclusive/Exclusive?</td>");
                                }
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Fund</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Other Fund</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Not sure of fund</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Co-payment</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Purpose of payment</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">NMDS Payment purpose</td>"); // NEW
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Other payment purpose</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Locality (suburb)</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Weekly household rent $</td>");
                                sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Fund (Legacy data)</td>");
                                sbHeaderList.AppendLine("</tr>");

                                // Add report data to report
                                if (sbReportList.Length > 0)
                                    sbHeaderList.AppendLine(sbReportList.ToString());

                                sbHeaderList.AppendLine("</table>");

                                byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                                string encodedData = System.Convert.ToBase64String(filename);
                                Entity Annotation = new Entity("annotation");
                                Annotation.Attributes["objectid"] = new EntityReference("new_palmgofinancials", varFinancialID);
                                Annotation.Attributes["objecttypecode"] = "new_palmgofinancials";
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
                        else if (varReport == "Client financial data")
                        {


                            // Get the report name and file name
                            varReportName = varBrokerage + " extract for " + cleanDate(varStartDate) + " to " + cleanDate(varEndDate);
                            varFileName = varBrokerage + " extract for " + cleanDate(varStartDate) + " to " + cleanDate(varEndDate) + ".xls";

                            sbHeaderList.Clear();
                            sbReportList.Clear();

                            //Clear variables.
                            //varFileName = "No Report"; // Report file name
                            //varReportName = "No Report"; // Name of report
                            dbPalmClientFinancialId = "";
                            dbEntryDate = "";
                            dbAmount = 0;
                            dbGst = 0;
                            dbPalmClientId = "";
                            dbClient = "";
                            dbWeekRent = "";
                            dbPalmClientSupportId = "";
                            dbBrokerage = "";
                            varSurname = "";
                            varFirstName = "";
                            varGender = "";

                            // Fetch statements for database
                            // Get the financial data for the period and brokerage chosen
                            dbFinancialList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclient'>
                                <attribute name='new_palmclientid' />
                                <attribute name='new_address' />
                                <attribute name='new_firstname' />
                                <attribute name='new_surname' />
                                <attribute name='new_gender' />
                                <attribute name='new_indigenous' />
                                <attribute name='new_clientage' />
                                <link-entity name='new_palmclientsupport' to='new_palmclientid' from='new_client' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                    <link-entity name='new_palmclientfinancial' to='new_palmclientsupportid' from='new_supportperiod' link-type='inner'>
                                        <attribute name='new_entrydate' />
                                        <attribute name='new_amount' />
                                        <attribute name='new_gst' />
                                        <attribute name='new_assistance' />
                                        <attribute name='new_brokerage' />
                                        <link-entity name='new_palmsubrokerage' to='new_brokerage' from='new_palmsubrokerageid' link-type='outer'>
                                            <attribute name='new_shorthand' />
                                        </link-entity>
                                    </link-entity>
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_amount' operator='gt' value='0' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='eq' value='" + ownerLookup.Id + @"' />
                                </filter>
                              </entity>
                            </fetch> ";


                            // Get the fetch XML data and place in entity collection objects
                            result = _service.RetrieveMultiple(new FetchExpression(dbFinancialList));

                            // Loop through financial data
                            foreach (var c in result.Entities)
                            {


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

                                if (c.FormattedValues.Contains("new_clientage"))
                                    dbAge = c.FormattedValues["new_clientage"];
                                else if (c.Attributes.Contains("new_clientage"))
                                    dbAge = c.Attributes["new_clientage"].ToString();
                                else
                                    dbAge = "0"; //Bad that age is a string.

                                if (c.FormattedValues.Contains("new_gender"))
                                    dbGender = c.FormattedValues["new_gender"];
                                else if (c.Attributes.Contains("new_gender"))
                                    dbGender = c.Attributes["new_gender"].ToString();
                                else
                                    dbGender = "";

                                if (c.FormattedValues.Contains("new_indigenous"))
                                    dbIndigenous = c.FormattedValues["new_indigenous"];
                                else if (c.Attributes.Contains("new_indigenous"))
                                    dbIndigenous = c.Attributes["new_indigenous"].ToString();
                                else
                                    dbIndigenous = "";

                                if (c.FormattedValues.Contains("new_palmclientfinancial2.new_entrydate"))
                                    dbEntryDate = c.FormattedValues["new_palmclientfinancial2.new_entrydate"];
                                else if (c.Attributes.Contains("new_palmclientfinancial2.new_entrydate"))
                                    dbEntryDate = c.Attributes["new_palmclientfinancial2.new_entrydate"].ToString();
                                else
                                    dbEntryDate = "";

                                // Convert date from American format to Australian format
                                dbEntryDate = cleanDateAM(dbEntryDate);

                                if (c.FormattedValues.Contains("new_palmclientfinancial2.new_amount"))
                                    Double.TryParse(c.FormattedValues["new_palmclientfinancial2.new_amount"], out dbAmount);
                                else if (c.Attributes.Contains("new_palmclientfinancial2.new_amount"))
                                    Double.TryParse(c.Attributes["new_palmclientfinancial2.new_amount"].ToString(), out dbAmount);
                                else
                                    dbAmount = 0;

                                if (c.FormattedValues.Contains("new_palmclientfinancial2.new_gst"))
                                    Double.TryParse(c.FormattedValues["new_palmclientfinancial2.new_gst"], out dbGst);
                                else if (c.Attributes.Contains("new_palmclientfinancial2.new_gst"))
                                    Double.TryParse(c.Attributes["new_palmclientfinancial2.new_gst"].ToString(), out dbGst);
                                else
                                    dbGst = 0;

                                if (c.FormattedValues.Contains("new_palmclientfinancial2.new_assistance"))
                                    dbAssistance = c.FormattedValues["new_palmclientfinancial2.new_assistance"];
                                else if (c.Attributes.Contains("new_palmclientfinancial2.new_assistance"))
                                    dbAssistance = c.Attributes["new_palmclientfinancial2.new_assistance"].ToString();
                                else
                                    dbAssistance = "";

                                
                                if (c.FormattedValues.Contains("new_palmclientfinancial2.new_brokerage"))
                                    dbBrokerage = c.FormattedValues["new_palmclientfinancial2.new_brokerage"];
                                else if (c.Attributes.Contains("new_palmclientfinancial2.new_brokerage"))
                                    dbBrokerage = c.Attributes["new_palmclientfinancial2.new_brokerage"].ToString();
                                else
                                    dbBrokerage = "";


                                // Append data to report
                                sbReportList.Append("<tr>\r\n<td>" + dbClient + "</td>\r\n<td>" + dbFirstName + "</td>\r\n<td>" + dbSurname + "</td>\r\n<td>" + dbAge + "</td>\r\n<td>" + dbGender + "</td>\r\n<td>" + dbIndigenous + "</td>\r\n<td>" + dbEntryDate + "</td>\r\n<td>" + dbAmount + "</td>\r\n<td>" + dbGst + "</td>\r\n<td>" + dbAssistance + "</td>\r\n<td>");

                                sbReportList.Append(dbBrokerage);

                                sbReportList.AppendLine("</td>\r\n</tr>");

                            } // client loop        

                            //Header part of the Report
                            sbHeaderList.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                            sbHeaderList.AppendLine("<head>");
                            sbHeaderList.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                            sbHeaderList.AppendLine("<!--[if gte mso 9]>");
                            sbHeaderList.AppendLine("<xml>");
                            sbHeaderList.AppendLine("<x:ExcelWorkbook>");
                            sbHeaderList.AppendLine("<x:ExcelWorksheets>");
                            sbHeaderList.AppendLine("<x:ExcelWorksheet>");

                            //this line names the worksheet
                            sbHeaderList.AppendLine("<x:Name>Financial Data</x:Name>");

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
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Client Num</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">First Name</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Last Name</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Age</td>"); // LOOK AT
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Gender</td>"); // NEW
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">ATSI</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Financial Date</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Amount $</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">GST $</td>"); // NEW
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Assistance</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Brokerage</td>");
                            sbHeaderList.AppendLine("</tr>");

                            // Add report data to report
                            if (sbReportList.Length > 0)
                                sbHeaderList.AppendLine(sbReportList.ToString());

                            sbHeaderList.AppendLine("</table>");

                            byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                            string encodedData = System.Convert.ToBase64String(filename);
                            Entity Annotation = new Entity("annotation");
                            Annotation.Attributes["objectid"] = new EntityReference("new_palmgofinancials", varFinancialID);
                            Annotation.Attributes["objecttypecode"] = "new_palmgofinancials";
                            Annotation.Attributes["subject"] = "Report";
                            Annotation.Attributes["documentbody"] = encodedData;
                            Annotation.Attributes["mimetype"] = @"application/msexcel";
                            Annotation.Attributes["notetext"] = varReportName;
                            Annotation.Attributes["filename"] = varFileName;
                            _service.Create(Annotation);

                            sbHeaderList.Clear();
                            sbReportList.Clear();
                        }
                        else if (varReport == "Kypera Spend Report")
                        {


                            
                            sbHeaderList.Clear();
                            sbReportList.Clear();

                            //Clear variables.
                            //varFileName = "No Report"; // Report file name
                            //varReportName = "No Report"; // Name of report
                            dbPalmClientFinancialId = "";
                            dbPaidDate = "";
                            dbPaidBy = "";
                            dbAmount = 0;
                            dbGst = 0;
                            dbBrokerage = "";
                            varSurname = "";
                            varFirstName = "";
                            varGSTUpDown = "";
                            dbBankAcc = "";
                            dbAssistanceCode = "";
                            dbCheque = "";

                            string bankAccountName = "";

                            switch (varBankAccount.Value)
                            {
                                case 100000000:
                                    bankAccountName = "ER-11152";
                                    break;
                                case 100000001:
                                    bankAccountName = "Flexcare-11153";
                                    break;
                                case 100000002:
                                    bankAccountName = "Geelong ER - 11125";
                                    break;
                                case 100000003:
                                    bankAccountName = "HEF - 11151";
                                    break;
                                case 100000004:
                                    bankAccountName = "Mallee - 11154";
                                    break;
                                case 100000005:
                                    bankAccountName = "Metro HEF - 11155";
                                    break;
                            }

                            // Get the report name and file name
                            varReportName = bankAccountName + " kypera spend report for " + cleanDate(varStartDate) + " to " + cleanDate(varEndDate);
                            varFileName = bankAccountName + " kypera spend report for " + cleanDate(varStartDate) + " to " + cleanDate(varEndDate) + ".xls";


                            // Fetch statements for database
                            // Get the financial data for the period and brokerage chosen
                            dbFinancialList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                                <entity name='new_palmclientfinancial'>
                                    <attribute name='new_firstnamenc' />
                                    <attribute name='new_gst' />
                                    <attribute name='new_payee' />
                                    <attribute name='new_cheque' />
                                    <attribute name='new_assistance'/>
                                    <attribute name='new_amount'/>
                                    <attribute name='new_datepaid'/>
                                    <attribute name='new_brokerage'/>
                                    <attribute name='new_surnamenc'/>
                                    <link-entity name='new_palmsuassistance' from='new_palmsuassistanceid' to='new_assistance'>
                                        <attribute name='new_code'/>
                                    </link-entity>
                                    <link-entity name='new_palmsubrokerage' from='new_palmsubrokerageid' to='new_brokerage'>
                                        <attribute name='new_progid'/>
                                        <attribute name='new_bankacc'/>
                                    </link-entity>
                                    <link-entity name='new_palmclientcasenotes' from='new_palmclientcasenotesid' to='new_casenote' link-type='outer'>
                                        <link-entity name='new_palmclientsupport' from='new_palmclientsupportid' to='new_supportperiod' link-type='outer'>
                                            <attribute name='new_surname'/>
                                            <attribute name='new_firstname'/>
                                        </link-entity>
                                    </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientfinancial' attribute='new_datepaid' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_datepaid' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_amount' operator='gt' value='0' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_lockedfin' operator='eq' value='0' />
                                    <condition entityname='new_palmsubrokerage' attribute='new_bankacc' operator='eq' value='" + bankAccountName + @"' />
                                    <filter type='or'>
                                        <condition entityname='new_palmclientfinancial' attribute='new_paidby' operator='eq' value='100000000' />
                                        <condition entityname='new_palmclientfinancial' attribute='new_paidby' operator='eq' value='100000005' />
                                    </filter>
                                </filter>
                              </entity>
                            </fetch> ";

                            // Get the fetch XML data and place in entity collection objects
                            result = _service.RetrieveMultiple(new FetchExpression(dbFinancialList));

                            int count = 0;
                            // Loop through financial data
                            foreach (var c in result.Entities)
                            {
                                count++;
                                varGSTUpDown = "";

                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (c.FormattedValues.Contains("new_datepaid"))
                                    dbPaidDate = c.FormattedValues["new_datepaid"];
                                else if (c.Attributes.Contains("new_datepaid"))
                                    dbPaidDate = c.Attributes["new_datepaid"].ToString();
                                else
                                    dbPaidDate = "";

                                // Convert date from American format to Australian format
                                dbEntryDate = cleanDateAM(dbEntryDate);

                                if (c.FormattedValues.Contains("new_cheque"))
                                    dbCheque = c.FormattedValues["new_cheque"];
                                else if (c.Attributes.Contains("new_cheque"))
                                    dbCheque = c.GetAttributeValue<AliasedValue>("new_cheque").Value.ToString();
                                else
                                    dbCheque = "";

                                if (c.FormattedValues.Contains("new_payee"))
                                    dbPayee = c.FormattedValues["new_payee"];
                                else if (c.Attributes.Contains("new_payee"))
                                    dbPayee = c.Attributes["new_payee"].ToString();
                                else
                                    dbPayee = "";

                                if (c.FormattedValues.Contains("new_amount"))
                                    Double.TryParse(c.FormattedValues["new_amount"], out dbAmount);
                                else if (c.Attributes.Contains("new_amount"))
                                    Double.TryParse(c.Attributes["new_amount"].ToString(), out dbAmount);
                                else
                                    dbAmount = 0;

                                if (c.FormattedValues.Contains("new_gst"))
                                    Double.TryParse(c.FormattedValues["new_gst"], out dbGst);
                                else if (c.Attributes.Contains("new_gst"))
                                    Double.TryParse(c.Attributes["new_gst"].ToString(), out dbGst);
                                else
                                    dbGst = 0;

                                if (c.FormattedValues.Contains("new_firstnamenc"))
                                    dbFirstName = c.FormattedValues["new_firstnamenc"];
                                else if (c.Attributes.Contains("new_firstnamenc"))
                                    dbFirstName = c.Attributes["new_firstnamenc"].ToString();
                                else
                                    dbFirstName = "";

                                if (c.FormattedValues.Contains("new_surnamenc"))
                                    dbSurname = c.FormattedValues["new_surnamenc"];
                                else if (c.Attributes.Contains("new_surnamenc"))
                                    dbSurname = c.Attributes["new_surnamenc"].ToString();
                                else
                                    dbSurname = "";

                                //WORK OUT GST CODE
                                if (dbGst == 0)
                                {
                                    //varGSTUpDown += " ...Gst was blank."; //Gst was blank
                                }
                                else if (Math.Abs(dbGst - Math.Round(dbAmount / 1.1 / 10, 2)) > 0.015)
                                { //Gst is not included.
                                    varGSTUpDown += " ...Gst isn't 10% of amount: " + Math.Round((dbAmount / 1.1) / 10, 2);
                                }

                                //Assistance Fields
                                if (c.FormattedValues.Contains("new_palmsuassistance1.new_code"))
                                    dbAssistance = c.FormattedValues["new_palmsuassistance1.new_code"];
                                else if (c.Attributes.Contains("new_palmsuassistance1.new_code"))
                                    dbAssistance = c.GetAttributeValue<AliasedValue>("new_palmsuassistance1.new_code").Value.ToString();
                                else
                                    dbAssistance = "";

                                switch (dbAssistance)
                                {
                                    case "55210":
                                    case "55220":
                                    case "55060":
                                        dbGstCode = "G14";
                                        varGSTUpDown = "";
                                        break;
                                    default:
                                        dbGstCode = "G11";
                                        break;
                                }

                                //Brokerage fields
                                if (c.FormattedValues.Contains("new_palmsubrokerage2.new_bankacc"))
                                    dbBankAcc = c.FormattedValues["new_palmsubrokerage2.new_bankacc"];
                                else if (c.Attributes.Contains("new_palmsubrokerage2.new_bankacc"))
                                    dbBankAcc = c.GetAttributeValue<AliasedValue>("new_palmsubrokerage2.new_bankacc").Value.ToString();
                                else
                                    dbBankAcc = "";

                                if (dbBankAcc != "")
                                {
                                    dbBankAcc = cleanString(dbBankAcc, "bankAccount");
                                    var accountIndex = dbBankAcc.IndexOf('-'); //Get where '-' symbol appears
                                    if (accountIndex <= 0)
                                        accountIndex = 0;
                                    dbBankAcc = dbBankAcc.Substring(accountIndex, dbBankAcc.Length); //Shorten to just
                                }

                                if (c.FormattedValues.Contains("new_palmsubrokerage2.new_progid"))
                                    dbProgId = c.FormattedValues["new_palmsubrokerage2.new_progid"];
                                else if (c.Attributes.Contains("new_palmsubrokerage2.new_progid"))
                                    dbProgId = c.GetAttributeValue<AliasedValue>("new_palmsubrokerage2.new_progid").Value.ToString();
                                else
                                    dbProgId = "";

                                //Support period fields
                                if (dbFirstName == "" || dbFirstName == null)
                                {
                                    if (c.FormattedValues.Contains("new_palmclientsupport4.new_firstname"))
                                        dbFirstName = c.FormattedValues["new_palmclientsupport4.new_firstname"];
                                    else if (c.Attributes.Contains("new_palmclientsupport4.new_firstname"))
                                        dbFirstName = c.GetAttributeValue<AliasedValue>("new_palmclientsupport4.new_firstname").Value.ToString();
                                    else
                                        dbFirstName = "";

                                    if (c.FormattedValues.Contains("new_palmclientsupport4.new_surname"))
                                        dbSurname = c.FormattedValues["new_palmclientsupport4.new_surname"];
                                    else if (c.Attributes.Contains("new_palmclientsupport4.new_surname"))
                                        dbSurname = c.GetAttributeValue<AliasedValue>("new_palmclientsupport4.new_surname").Value.ToString();
                                    else
                                        dbSurname = "";
                                }

                                // Append data to report
                                sbReportList.Append("<tr>\r\n<td>" + count + "</td>\r\n<td>1</td>\r\n<td>" + dbPaidDate + "</td>\r\n<td>1" + dbBankAcc + "</td>\r\n<td>" + dbPayee + "</td>\r\n<td>" + dbFirstName + " " + dbSurname + "</td>\r\n<td>" + dbPayee + "</td>\r\n<td>" + dbFirstName + " " + dbSurname + "</td>\r\n<td></td>\r\n<td>" + dbCheque + "</td>\r\n<td>-$" + dbAmount + "</td>\r\n<td>N-T</td>\r\n<td>$0.00</td>\r\n<td></td>\r\n<td></td>\r\n<td></td>\r\n</tr>\r\n");
                                sbReportList.Append("<tr>\r\n<td>" + count + "</td>\r\n<td>2</td>\r\n<td>" + dbPaidDate + "</td>\r\n<td>1-" + dbAssistance + "</td>\r\n<td>" + dbPayee + "</td>\r\n<td>" + dbFirstName + " " + dbSurname + "</td>\r\n<td>" + dbPayee + "</td>\r\n<td>" + dbFirstName + " " + dbSurname + "</td>\r\n<td></td>\r\n<td></td>\r\n<td>$" + dbAmount + "</td>\r\n<td>" + dbGstCode + "</td>\r\n<td>$" + dbGst + varGSTUpDown + "</td>\r\n<td>2</td>\r\n<td>PBR</td>\r\n<td>" + dbProgId + "</td>\r\n</tr>\r\n");
                                
                            } // client loop        

                            //Header part of the Report
                            sbHeaderList.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                            sbHeaderList.AppendLine("<head>");
                            sbHeaderList.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                            sbHeaderList.AppendLine("<!--[if gte mso 9]>");
                            sbHeaderList.AppendLine("<xml>");
                            sbHeaderList.AppendLine("<x:ExcelWorkbook>");
                            sbHeaderList.AppendLine("<x:ExcelWorksheets>");
                            sbHeaderList.AppendLine("<x:ExcelWorksheet>");

                            //this line names the worksheet
                            sbHeaderList.AppendLine("<x:Name>Financial Data</x:Name>");

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
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">TransactionNo</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">LineNo</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">TransactionDate</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">FullAccount</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">TransactionDesc</td>"); 
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">InvoiceRef</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">PaymentRef</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">ExternalRef</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">InternalRef</td>"); 
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">ChequeNo</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">GrossAmount</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">TaxCode</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">TaxAmount</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">District</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Scheme</td>");
                            sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Block</td>");
                            sbHeaderList.AppendLine("</tr>");
                            // Add report data to report
                            if (sbReportList.Length > 0)
                                sbHeaderList.AppendLine(sbReportList.ToString());

                            sbHeaderList.AppendLine("</table>");

                            byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                            string encodedData = System.Convert.ToBase64String(filename);
                            Entity Annotation = new Entity("annotation");
                            Annotation.Attributes["objectid"] = new EntityReference("new_palmgofinancials", varFinancialID);
                            Annotation.Attributes["objecttypecode"] = "new_palmgofinancials";
                            Annotation.Attributes["subject"] = "Report";
                            Annotation.Attributes["documentbody"] = encodedData;
                            Annotation.Attributes["mimetype"] = @"application/msexcel";
                            Annotation.Attributes["notetext"] = varReportName;
                            Annotation.Attributes["filename"] = varFileName;
                            _service.Create(Annotation);

                            sbHeaderList.Clear();
                            sbReportList.Clear();
                        }
                        /*
                        //varTest += sbHeaderList.ToString();

                        // Create note against current Palm Go Financial record and add attachment
                        byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                        string encodedData = System.Convert.ToBase64String(filename);
                        Entity Annotation = new Entity("annotation");
                        Annotation.Attributes["objectid"] = new EntityReference("new_palmgofinancials", varFinancialID);
                        Annotation.Attributes["objecttypecode"] = "new_palmgofinancials";
                        Annotation.Attributes["subject"] = "Report";
                        Annotation.Attributes["documentbody"] = encodedData;
                        Annotation.Attributes["mimetype"] = @"application/msexcel";
                        Annotation.Attributes["notetext"] = varReportName;
                        Annotation.Attributes["filename"] = varFileName;
                        _service.Create(Annotation);
                        
                         */

                        // Add the second report if relevant
                        if (varReport2 == true)
                        {
                            byte[] filename2 = Encoding.ASCII.GetBytes(sbReport2List.ToString());
                            string encodedData2 = System.Convert.ToBase64String(filename2);
                            Entity Annotation2 = new Entity("annotation");
                            Annotation2.Attributes["objectid"] = new EntityReference("new_palmgofinancials", varFinancialID);
                            Annotation2.Attributes["objecttypecode"] = "new_palmgofinancials";
                            Annotation2.Attributes["subject"] = "Report";
                            Annotation2.Attributes["documentbody"] = encodedData2;
                            Annotation2.Attributes["mimetype"] = @"application/msexcel";
                            Annotation2.Attributes["notetext"] = varReportName2;
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
                varCharAllowed = "abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ-'"; //characters allowed
            else if (thetype == "username")
                varCharAllowed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"; //characters allowed
            else if (thetype == "mailbox")
                varCharAllowed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890-."; //characters allowed
            else if (thetype == "slk")
                varCharAllowed = "ABCDEFGHIJKLMNOPQRSTUVWXYZ29"; //characters allowed
            else if (thetype == "letter")
                varCharAllowed = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"; //characters allowed
            else if (thetype == "search")
                varCharAllowed = "abcdefghijklmnopqrstuvwxyz1234567890."; //characters allowed
            else if (thetype == "palm")
                varCharAllowed = "PLMplm1234567890"; //characters allowed
            else if (thetype == "voucher")
                varCharAllowed = "abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ-_()1234567890#"; //characters allowed
            else if (thetype == "bankAccount")
                varCharAllowed = "1234567890-."; //characters allowed
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
}

