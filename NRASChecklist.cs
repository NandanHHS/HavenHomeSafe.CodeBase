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
    public class nrasChecklist : IPlugin
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

                // Verify that the target entity represents the NRAS Checklist entity
                // If not, this plug-in was not registered correctly.
                // if (entity.LogicalName != "new_nraschecklist")
                //    return;

                try
                {
                    // Fixes:
                    // Remove ampersand from comments
                    // Compare values to Lower

                    string varDescription = ""; // Description
                    bool varCreateExtract = false; // Whether to generate extract
                    string varDwellingName = ""; // Dwelling name
                    Guid varDwellingID = new Guid(); // GUID for dwelling
                    int varCheckInt = 0; // Used to parse integers
                    double varCheckDouble = 0; // Used to parse doubles
                    DateTime varCheckDate = new DateTime(); // Used to parse dates

                    StringBuilder sbHeaderList = new StringBuilder(); // String builder for header data
                    StringBuilder sbClientList = new StringBuilder(); // String builder for data rows
                    string varFileName = ""; // File name

                    string varTest = ""; // Debug

                    // Only do this if the entity is the NRAS checklist entity
                    if (entity.LogicalName == "new_nraschecklist")
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

                        varDwellingID = entity.Id; // Get the dwelling id

                        // Get the associated data from the dwelling
                        EntityReference ownerLookup = (EntityReference)entity.Attributes["new_dwellingid"];
                        var actualOwningUnit = _service.Retrieve(ownerLookup.LogicalName, ownerLookup.Id, new ColumnSet(true));
                        varDwellingName = actualOwningUnit["new_dewllid"].ToString();

                        // Fetch statements for database
                        // Get the data from the NRAS property table based on the property id
                        string dbDwellingList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_nrasproperty'>
                                <attribute name='new_dewllid' />
                                <attribute name='new_address' />
                                <attribute name='new_propmgr' />
                                <attribute name='new_startdate' />
                                <filter type='and'>
                                    <condition entityname='new_nrasproperty' attribute='new_nraspropertyid' operator='eq' value='" + ownerLookup.Id + @"' />
                                </filter>
                                <order entityname='new_nrasproperty' attribute='new_dewllid' />
                              </entity>
                            </fetch> ";

                        // Get the data from the NRAS tenancy table based on the property id
                        string dbTenancyList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_nrastenancy'>
                                <attribute name='new_dwellid' />
                                <attribute name='new_tenant' />
                                <attribute name='new_scenario' />
                                <attribute name='new_tdadate' />
                                <attribute name='new_vacatedate' />
                                <attribute name='new_leasestartdate' />
                                <attribute name='new_leaseexpirydate' />
                                <attribute name='new_marketrent' />
                                <attribute name='new_maxweekrent' />
                                <attribute name='new_maxmonthlyrent' />
                                <attribute name='new_currnrasrent' />
                                <attribute name='new_income' />
                                <attribute name='new_leaseprov' />
                                <filter type='and'>
                                    <condition entityname='new_nrastenancy' attribute='new_dwellid' operator='eq' value='" + ownerLookup.Id + @"' />
                                </filter>
                                <order entityname='new_nrastenancy' attribute='new_leasestartdate' />
                              </entity>
                            </fetch> ";

                        // Variables to hold database data
                        string dbDwellId = "";
                        string dbAddress = "";
                        string dbPropMgr = "";
                        string dbStartDate = "";

                        string dbTenant = "";
                        string dbScenario = "";
                        string dbTdaDate = "";
                        string dbVacateDate = "";
                        string dbLeaseStartDate = "";
                        string dbLeaseExpiryDate = "";
                        string dbMarketRent = "";
                        string dbMaxWeekRent = "";
                        string dbMaxMonthlyRent = "";
                        string dbCurrNrasRent = "";
                        string dbIncome = "";
                        string dbLeaseProv = "";

                        string dbVacateDate2 = "";
                        bool varFirstRecord = false; // Whether this is the first record
                        string varCurrYear = ""; // Current year
                        string varCurrYear2 = ""; // Current year 2

                        // Get the fetch XML data and place in entity collection objects
                        EntityCollection result = _service.RetrieveMultiple(new FetchExpression(dbDwellingList));
                        EntityCollection result2 = _service.RetrieveMultiple(new FetchExpression(dbTenancyList));

                        // Loop through the property data
                        foreach (var c in result.Entities)
                        {
                            //varTest += "in";

                            // Process the data as follows:
                            // If there is a formatted value for the field, use it
                            // Otherwise if there is a literal value for the field, use it
                            // Otherwise the value wasn't returned so set as nothing
                            if (c.FormattedValues.Contains("new_dewllid"))
                                dbDwellId = c.FormattedValues["new_dewllid"];
                            else if(c.Attributes.Contains("new_dewllid"))
                                dbDwellId = c.Attributes["new_dewllid"].ToString();
                            else
                                dbDwellId = "";

                            if (c.FormattedValues.Contains("new_address"))
                                dbAddress = c.FormattedValues["new_address"];
                            else if (c.Attributes.Contains("new_address"))
                                dbAddress = c.Attributes["new_address"].ToString();
                            else
                                dbAddress = "";

                            if (String.IsNullOrEmpty(dbAddress) == false)
                                dbAddress = dbAddress.Replace(",", "");

                            if (c.FormattedValues.Contains("new_propmgr"))
                                dbPropMgr = c.FormattedValues["new_propmgr"];
                            else if (c.Attributes.Contains("new_propmgr"))
                                dbPropMgr = c.Attributes["new_propmgr"].ToString();
                            else
                                dbPropMgr = "";

                            if (c.FormattedValues.Contains("new_startdate"))
                                dbStartDate = getDate(c.FormattedValues["new_startdate"]);
                            else if (c.Attributes.Contains("new_startdate"))
                                dbStartDate = getDate(c.Attributes["new_startdate"].ToString());
                            else
                                dbStartDate = "";
                        }

                        // Loop through the tenancy data
                        foreach (var d in result2.Entities)
                        {
                            // The first record returns all requird values, other records are for historic vacate dates
                            if (varFirstRecord == false)
                            {
                                // Process the data as follows:
                                // If there is a formatted value for the field, use it
                                // Otherwise if there is a literal value for the field, use it
                                // Otherwise the value wasn't returned so set as nothing
                                if (d.FormattedValues.Contains("new_tenant"))
                                    dbTenant = d.FormattedValues["new_tenant"];
                                else if (d.Attributes.Contains("new_tenant"))
                                    dbTenant = d.Attributes["new_tenant"].ToString();
                                else
                                    dbTenant = "";

                                if (d.FormattedValues.Contains("new_scenario"))
                                    dbScenario = d.FormattedValues["new_scenario"];
                                else if (d.Attributes.Contains("new_scenario"))
                                    dbScenario = d.Attributes["new_scenario"].ToString();
                                else
                                    dbScenario = "";

                                if (d.FormattedValues.Contains("new_tdadate"))
                                    dbTdaDate = getDate(d.FormattedValues["new_tdadate"]);
                                else if (d.Attributes.Contains("new_tdadate"))
                                    dbTdaDate = getDate(d.Attributes["new_tdadate"].ToString());
                                else
                                    dbTdaDate = "";

                                if (d.FormattedValues.Contains("new_vacatedate"))
                                    dbVacateDate = getDate(d.FormattedValues["new_vacatedate"]);
                                else if (d.Attributes.Contains("new_vacatedate"))
                                    dbVacateDate = getDate(d.Attributes["new_vacatedate"].ToString());
                                else
                                    dbVacateDate = "";

                                if (d.FormattedValues.Contains("new_leasestartdate"))
                                    dbLeaseStartDate = getDate(d.FormattedValues["new_leasestartdate"]);
                                else if (d.Attributes.Contains("new_leasestartdate"))
                                    dbLeaseStartDate = getDate(d.Attributes["new_leasestartdate"].ToString());
                                else
                                    dbLeaseStartDate = "";

                                if (d.FormattedValues.Contains("new_leaseexpirydate"))
                                    dbLeaseExpiryDate = getDate(d.FormattedValues["new_leaseexpirydate"]);
                                else if (d.Attributes.Contains("new_leaseexpirydate"))
                                    dbLeaseExpiryDate = getDate(d.Attributes["new_leaseexpirydate"].ToString());
                                else
                                    dbLeaseExpiryDate = "";

                                if (d.FormattedValues.Contains("new_marketrent"))
                                    dbMarketRent = d.FormattedValues["new_marketrent"];
                                else if (d.Attributes.Contains("new_marketrent"))
                                    dbMarketRent = d.Attributes["new_marketrent"].ToString();
                                else
                                    dbMarketRent = "";

                                if (d.FormattedValues.Contains("new_maxweekrent"))
                                    dbMaxWeekRent = d.FormattedValues["new_maxweekrent"];
                                else if (d.Attributes.Contains("new_maxweekrent"))
                                    dbMaxWeekRent = d.Attributes["new_maxweekrent"].ToString();
                                else
                                    dbMaxWeekRent = "";

                                if (d.FormattedValues.Contains("new_maxmonthlyrent"))
                                    dbMaxMonthlyRent = d.FormattedValues["new_maxmonthlyrent"];
                                else if (d.Attributes.Contains("new_maxmonthlyrent"))
                                    dbMaxMonthlyRent = d.Attributes["new_maxmonthlyrent"].ToString();
                                else
                                    dbMaxMonthlyRent = "";

                                if (d.FormattedValues.Contains("new_currnrasrent"))
                                    dbCurrNrasRent = d.FormattedValues["new_currnrasrent"];
                                else if (d.Attributes.Contains("new_currnrasrent"))
                                    dbCurrNrasRent = d.Attributes["new_currnrasrent"].ToString();
                                else
                                    dbCurrNrasRent = "";

                                if (d.FormattedValues.Contains("new_income"))
                                    dbIncome = d.FormattedValues["new_income"];
                                else if (d.Attributes.Contains("new_income"))
                                    dbIncome = d.Attributes["new_income"].ToString();
                                else
                                    dbIncome = "";

                                if (d.FormattedValues.Contains("new_leaseprov"))
                                    dbLeaseProv = d.FormattedValues["new_leaseprov"];
                                else if (d.Attributes.Contains("new_leaseprov"))
                                    dbLeaseProv = d.Attributes["new_leaseprov"].ToString();
                                else
                                    dbLeaseProv = "";
                            }
                            else
                            {
                                if (d.FormattedValues.Contains("new_vacatedate"))
                                    dbVacateDate2 = getDate(d.FormattedValues["new_vacatedate"]);
                                else if (d.Attributes.Contains("new_vacatedate"))
                                    dbVacateDate2 = getDate(d.Attributes["new_vacatedate"].ToString());
                                else
                                    dbVacateDate2 = "";
                                break;
                            }

                            varFirstRecord = true; // flag as first record processed
                        }

                        // Get start and end date
                        if (DateTime.Now.Month <= 6)
                        {
                            varCurrYear = (DateTime.Now.Year - 1) + "";
                            varCurrYear2 = (DateTime.Now.Year) + "";
                        }
                        else
                        {
                            varCurrYear = (DateTime.Now.Year) + "";
                            varCurrYear2 = (DateTime.Now.Year + 1) + "";
                        }

                        //Header part of the Checklist
                        sbHeaderList.Append("Heading,");
                        sbHeaderList.AppendLine("HHS Audit Checklist - NRAS " + varCurrYear.Substring(2,2) + " / " + varCurrYear2.Substring(2, 2) + "");
                        sbHeaderList.Append("Dwell Id,");
                        sbHeaderList.AppendLine(dbDwellId + "");
                        sbHeaderList.Append("Address,");
                        sbHeaderList.AppendLine(dbAddress + "");
                        sbHeaderList.Append("Property Manager,");
                        sbHeaderList.AppendLine(dbPropMgr + "");

                        sbHeaderList.Append("Start Date,");

                        if (String.IsNullOrEmpty(dbStartDate) == false)
                            sbHeaderList.AppendLine(dbStartDate + "");
                        else
                            sbHeaderList.AppendLine(" ");

                        sbHeaderList.Append("Tenant,");
                        sbHeaderList.AppendLine(dbTenant + " ");
                        sbHeaderList.Append("Scenario,");
                        sbHeaderList.AppendLine(dbScenario + " ");
                        sbHeaderList.Append("TDA,");

                        if (String.IsNullOrEmpty(dbTdaDate) == false)
                            sbHeaderList.AppendLine("Yes");
                        else
                            sbHeaderList.AppendLine("No");

                        sbHeaderList.Append("Vacate Date,");

                        if (String.IsNullOrEmpty(dbVacateDate) == false)
                            sbHeaderList.AppendLine(dbVacateDate + "");
                        else
                            sbHeaderList.AppendLine(" ");

                        sbHeaderList.Append("Lease Start Date,");

                        if (String.IsNullOrEmpty(dbLeaseStartDate) == false)
                            sbHeaderList.AppendLine(dbLeaseStartDate + "");
                        else
                            sbHeaderList.AppendLine(" ");

                        sbHeaderList.Append("Lease Expiry Date,");

                        if (String.IsNullOrEmpty(dbLeaseExpiryDate) == false)
                            sbHeaderList.AppendLine(dbLeaseExpiryDate + "");
                        else
                            sbHeaderList.AppendLine(" ");

                        sbHeaderList.Append("Market Rent,");
                        sbHeaderList.AppendLine(dbMarketRent + "");
                        sbHeaderList.Append("Max Week Rent,");
                        sbHeaderList.AppendLine(dbMaxWeekRent + "");
                        sbHeaderList.Append("Monthly Rent,");
                        sbHeaderList.AppendLine(dbMaxMonthlyRent + "");
                        sbHeaderList.Append("Curr NRAS Rent,");
                        sbHeaderList.AppendLine(dbCurrNrasRent + "");
                        sbHeaderList.Append("Income,");
                        sbHeaderList.AppendLine(dbIncome + "");
                        sbHeaderList.Append("Lease Prov,");
                        sbHeaderList.AppendLine(dbLeaseProv + "");
                        sbHeaderList.Append("Previous Vacate Date,");

                        if (String.IsNullOrEmpty(dbVacateDate2) == false)
                            sbHeaderList.AppendLine(dbVacateDate2 + "");
                        else
                            sbHeaderList.AppendLine(" ");

                        sbHeaderList.Append("Notes,");

                        if (String.IsNullOrEmpty(dbVacateDate2) == false)
                            sbHeaderList.Append("Previous tenants vacated on " + dbVacateDate2 + " ");
                        else
                            sbHeaderList.Append("No previous tenant information ");

                        if (dbLeaseProv == "Yes")
                            sbHeaderList.Append("- ledger provided. ");
                        else
                            sbHeaderList.Append("- ledger not provided. ");

                        if (String.IsNullOrEmpty(dbVacateDate2) == false)
                            sbHeaderList.Append("Vacancy period " + cleanDate(Convert.ToDateTime(dbVacateDate2).AddDays(+1)) + " to " + cleanDate(Convert.ToDateTime(dbLeaseStartDate).AddDays(-1)) + " (" + ((Convert.ToDateTime(dbLeaseStartDate).AddDays(-1) - Convert.ToDateTime(dbVacateDate2).AddDays(+1)).TotalDays) + " days).");

                        sbHeaderList.Append("TDA for " + varCurrYear + "/" + varCurrYear2.Substring(2, 2) + " ");

                        if (String.IsNullOrEmpty(dbTdaDate) == false)
                            sbHeaderList.AppendLine("submitted on" + dbTdaDate + ". ");
                        else
                            sbHeaderList.AppendLine("not submitted. ");

                        // Create the file name
                        varFileName = "NRAS " + varCurrYear.Substring(2, 2) + " / " + varCurrYear2.Substring(2, 2) + " compliance Checklist for " + dbDwellId + ".csv";

                        // Create note against current NRAS Checklist record and add attachment
                        byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                        string encodedData = System.Convert.ToBase64String(filename);
                        Entity Annotation = new Entity("annotation");
                        Annotation.Attributes["objectid"] = new EntityReference("new_nraschecklist", varDwellingID);
                        Annotation.Attributes["objecttypecode"] = "new_nraschecklist";
                        Annotation.Attributes["subject"] = "NRAS Checklist";
                        Annotation.Attributes["documentbody"] = encodedData;
                        Annotation.Attributes["mimetype"] = @"text/csv";
                        Annotation.Attributes["notetext"] = "NRAS " + varCurrYear.Substring(2,2) + " / " + varCurrYear2.Substring(2, 2) + " compliance Checklist for " + dbDwellId;
                        Annotation.Attributes["filename"] = varFileName;
                        _service.Create(Annotation);

                        //throw new InvalidPluginExecutionException("This plugin is working. Dwelling is " + varDwellingName + "\r\n" + varTest);
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

        //Format: 1-Jan-1970
        public string cleanDate(DateTime getdate)
        {
            string clean = getdate.Day + "-" + getdate.ToString("MMM") + "-" + getdate.Year;
            return clean;
        }

        // Get Date
        public string getDate(string s)
        {
            DateTime sCheckDate = new DateTime();

            if (DateTime.TryParse(s, out sCheckDate))
                s = cleanDate(sCheckDate);
            else
                s = "";

            return s;
        }

        // Convert multiselect option set to values with asterisks around them for better string matching
        public string getMult(string s)
        {
            string[] myMult;
            string getVal = "";

            myMult = s.Split(';');
            foreach (string tst in myMult)
            {
                getVal += "*" + tst.Trim() + "*,";
            }

            return getVal;
        }
    }
}


