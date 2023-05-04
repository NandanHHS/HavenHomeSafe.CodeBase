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

namespace Microsoft.Crm.Sdk.Samples
{
	public class goPRAP: IPlugin
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
                // if (entity.LogicalName != "new_palmgoprap")
                //    return;

                try
                {
                    // Fixes:
                    // Remove ampersand from comments
                    // Compare values to Lower

                    string varDescription = ""; // Description from form
                    bool varCreateExtract = false; // Create extract from form
                    string varAgencyId = ""; // Agency Id from form
                    Guid varPrapID = new Guid(); // GUID for PRAP
                    StringBuilder sbHeaderList = new StringBuilder(); // String builder for header
                    StringBuilder sbReportList = new StringBuilder(); // String builder for report
                    StringBuilder sbErrorList = new StringBuilder(); // String builder for errors

                    string varFileName = ""; // File name for report
                    string varFileName2 = ""; // File name for report
                    DateTime varStartDate = new DateTime(); // Start date
                    DateTime varEndDate = new DateTime(); // End date
                    DateTime varStartDatePr = new DateTime(); // Print start date
                    DateTime varEndDatePr = new DateTime(); // Print end date

                    int varCheckInt = 0; // Used to parse integers
                    double varCheckDouble = 0; // Used to parse doubles
                    DateTime varCheckDate = new DateTime(); // Used to parse dates
                    EntityReference getEntity; // Object for Entity Reference
                    AliasedValue getAlias; // Object for Aliased Value

                    string varTest = ""; // Debug

                    // Only do this if the entity is the Palm Go PRAP entity
                    if (entity.LogicalName == "new_palmgoprap")
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

                        // Get info for current PRAP record
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

                        varPrapID = entity.Id; // Get GUID

                        // Get associated values from PRAP agency
                        EntityReference ownerLookup = (EntityReference)entity.Attributes["new_agencyid"];
                        varAgencyId += ownerLookup.Id.ToString() + ".\r\n";
                        varAgencyId += ((EntityReference)entity.Attributes["new_agencyid"]).Name + ".\r\n";
                        varAgencyId += ownerLookup.LogicalName + ".\r\n";

                        var actualOwningUnit = _service.Retrieve(ownerLookup.LogicalName, ownerLookup.Id, new ColumnSet(true));
                        varAgencyId = actualOwningUnit["new_agency"].ToString();

                        // Create file name
                        varFileName = varAgencyId + " Report " + varStartDatePr.ToString("MMM") + " " + varStartDatePr.Year + " ending " + varEndDatePr.ToString("MMM") + " " + varEndDatePr.Year + ".xls";

                        // Create errors file name
                        varFileName2 = varFileName.Replace(".xls", ".txt");
                        varFileName2 = "Errors for " + varFileName2;

                        // Fetch statements for database
                        // Get the financial data for the period and brokerage chosen
                        string dbFinancialList = @"
                           <fetch version='1.0' mapping='logical' distinct='false'>
                              <entity name='new_palmclient'>
                                <attribute name='new_palmclientid' />
                                <attribute name='new_address' />
                                <attribute name='new_firstname' />
                                <attribute name='new_surname' />
                                <attribute name='new_gender' />
                                <attribute name='new_dob' />
                                <attribute name='new_dobest' />
                                <attribute name='new_mdsslk' />
                                <attribute name='new_shorslk' />
                                <attribute name='new_weekrent' />
                                <link-entity name='new_palmclientsupport' to='new_palmclientid' from='new_client' link-type='inner'>
                                    <attribute name='new_palmclientsupportid' />
                                    <attribute name='new_doshor' />
                                    <attribute name='new_locality' />
                                    <attribute name='new_puhid' />
                                    <link-entity name='new_palmclientfinancial' to='new_palmclientsupportid' from='new_supportperiod' link-type='inner'>
                                        <attribute name='new_entrydate' />
                                        <attribute name='new_amount' />
                                        <attribute name='new_assistance' />
                                    </link-entity>
                                    <link-entity name='new_palmddllocality' to='new_locality' from='new_palmddllocalityid' link-type='outer'>
                                        <attribute name='new_postcode' />
                                        <attribute name='new_state' />
                                    </link-entity>
                                </link-entity>
                                <filter type='and'>
                                    <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + cleanDate(varStartDate) + @"' />
                                    <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='le' value='" + cleanDate(varEndDate) + @"' />
                                    <condition entityname='new_palmclientsupport' attribute='new_doshor' operator='eq' value='" + ownerLookup.Id + @"' />
                                </filter>
                              </entity>
                            </fetch> ";

                        // Database variables
                        string dbPalmClientId = "";
                        string dbClient = "";
                        string dbFirstName = "";
                        string dbSurname = "";
                        string dbGender = "";
                        string dbDob = "";
                        string dbDobEst = "";
                        string dbMdsSlk = "";
                        string dbShorSlk = "";
                        string dbWeekRent = "";
                        string dbPalmClientSupportId = "";
                        string dbDoShor = "";
                        string dbLocality = "";
                        string dbState = "";
                        string dbPostcode = "";
                        string dbPuhId = "";
                        string dbEntryDate = "";
                        string dbAmount = "";
                        string dbAssistance = "";

                        string varSLK = "";
                        string varSurname = "";
                        string varFirstName = "";
                        string varDob = "";
                        string varGender = "";

                        // Get the fetch XML data and place in entity collection objects
                        EntityCollection result = _service.RetrieveMultiple(new FetchExpression(dbFinancialList));

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

                            if (c.FormattedValues.Contains("new_dobest"))
                                dbDobEst = c.FormattedValues["new_dobest"];
                            else if (c.Attributes.Contains("new_dobest"))
                                dbDobEst = c.Attributes["new_dobest"].ToString();
                            else
                                dbDobEst = "";

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
                                dbPalmClientSupportId = c.Attributes["new_palmclientsupport1.new_palmclientsupportid"].ToString();
                            else
                                dbPalmClientSupportId = "";

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
                                dbPuhId = "";

                            if (c.FormattedValues.Contains("new_palmclientfinancial2.new_entrydate"))
                                dbEntryDate = c.FormattedValues["new_palmclientfinancial2.new_entrydate"];
                            else if (c.Attributes.Contains("new_palmclientfinancial2.new_entrydate"))
                                dbEntryDate = c.Attributes["new_palmclientfinancial2.new_entrydate"].ToString();
                            else
                                dbEntryDate = "";

                            dbEntryDate = cleanDateAM(dbEntryDate);

                            if (c.FormattedValues.Contains("new_palmclientfinancial2.new_amount"))
                                dbAmount = c.FormattedValues["new_palmclientfinancial2.new_amount"];
                            else if (c.Attributes.Contains("new_palmclientfinancial2.new_amount"))
                                dbAmount = c.Attributes["new_palmclientfinancial2.new_amount"].ToString();
                            else
                                dbAmount = "";

                            dbAmount = cleanString(dbAmount, "double");

                            Double.TryParse(dbAmount, out varCheckDouble);
                            dbAmount = varCheckDouble.ToString("C");

                            if (c.FormattedValues.Contains("new_palmclientfinancial2.new_assistance"))
                                dbAssistance = c.FormattedValues["new_palmclientfinancial2.new_assistance"];
                            else if (c.Attributes.Contains("new_palmclientfinancial2.new_assistance"))
                                dbAssistance = c.Attributes["new_palmclientfinancial2.new_assistance"].ToString();
                            else
                                dbAssistance = "";

                            // Create the SLK based on firstname, surname, gender and dob
                            varSurname = varSurname + "22222";
                            varSurname = varSurname.Substring(1, 1) + varSurname.Substring(2, 1) + varSurname.Substring(4, 1);

                            if (varSurname == "222")
                                varSurname = "999";

                            varFirstName = varFirstName + "22222";
                            varFirstName = varFirstName.Substring(1, 1) + varFirstName.Substring(2, 1);

                            if (varFirstName == "22")
                                varFirstName = "99";

                            //Get the gender code
                            if (dbGender == "Female")
                                varGender = "2";
                            else if (dbGender == "Male")
                                varGender = "1";
                            else if (dbGender == "Intersex / Indeterminate") //Indeterminate should really be it's own value of 3, but is part of the intersex option.
                                varGender = "4";
                            else
                                varGender = "9"; //Not stated

                            //Put dob into expected format
                            varDob = cleanDateS(varCheckDate);

                            //Actual Dob  -Need to work out correct expected output for this.
                            //varCheckDate = DateTime.Parse(dbDob);
                            //varDob = cleanDateS(varCheckDate);

                            //Get the statistical linkage key
                            varSLK = varSurname + varFirstName + varDob + varGender;

                            if (String.IsNullOrEmpty(dbPuhId) == true)
                                dbPuhId = dbPalmClientSupportId;

                            // Append data to report
                            sbReportList.Append("<tr>\r\n<td>&nbsp;</td>\r\n<td>" + dbDoShor + "</td>\r\n<td>" + dbPuhId.Replace("-","") + "</td>\r\n<td>" + dbClient + "</td>\r\n<td>" + varSLK + "</td>\r\n<td>" + dbShorSlk + "</td>\r\n<td>" + dbEntryDate + "</td>\r\n<td>" + dbAmount + "</td>\r\n<td>");

                            sbReportList.Append("PRAP");

                            sbReportList.AppendLine("</td>\r\n<td>&nbsp;</td>\r\n<td>&nbsp;</td>\r\n<td>&nbsp;</td>\r\n<td>" + dbAssistance + "</td>\r\n<td>&nbsp;</td>\r\n<td>" + dbLocality + "</td>\r\n<td>" + dbWeekRent + "</td>\r\n<td>&nbsp;</td>\r\n</tr>");

                        } // client loop        


                        //Header part of the PRAP extract
                        sbHeaderList.AppendLine("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                        sbHeaderList.AppendLine("<head>");
                        sbHeaderList.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html;charset=windows-1252\">");
                        sbHeaderList.AppendLine("<!--[if gte mso 9]>");
                        sbHeaderList.AppendLine("<xml>");
                        sbHeaderList.AppendLine("<x:ExcelWorkbook>");
                        sbHeaderList.AppendLine("<x:ExcelWorksheets>");
                        sbHeaderList.AppendLine("<x:ExcelWorksheet>");

                        //this line names the worksheet
                        sbHeaderList.AppendLine("<x:Name>" + varAgencyId + " Data</x:Name>");

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
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Workgroup</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">SHS Agency ID</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Presenting Unit Link Key</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Client ID</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">SLK</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Alpha Code</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Date</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Amount $</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Fund</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Other Fund</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Not sure of fund</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Is another agency providing additional co-payment?</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Payment purpose</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Other payment purpose</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">The locality (suburb) of the lease (new or existing)</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Weekly rent required for tenancy $</td>");
                        sbHeaderList.AppendLine("<td class=\"prBorder\" style=\"font-weight: bold; white-space: nowrap; border-bottom: 2px solid #000000; background-color: #EEEEEE;\">Fund (Legacy data)</td>");
                        sbHeaderList.AppendLine("</tr>");

                        // Add report data to report
                        if (sbReportList.Length > 0)
                            sbHeaderList.AppendLine(sbReportList.ToString());

                        sbHeaderList.AppendLine("</table>");

                        //varTest += sbHeaderList.ToString();

                        // Create note against current Palm Go PRAP record and add attachment
                        byte[] filename = Encoding.ASCII.GetBytes(sbHeaderList.ToString());
                        string encodedData = System.Convert.ToBase64String(filename);
                        Entity Annotation = new Entity("annotation");
                        Annotation.Attributes["objectid"] = new EntityReference("new_palmgoprap", varPrapID);
                        Annotation.Attributes["objecttypecode"] = "new_palmgoprap";
                        Annotation.Attributes["subject"] = "PRAP Extract";
                        Annotation.Attributes["documentbody"] = encodedData;
                        Annotation.Attributes["mimetype"] = @"application/msexcel";
                        Annotation.Attributes["notetext"] = "PRAP Extract for " + cleanDate(varStartDatePr) + " to " + cleanDate(varEndDatePr);
                        Annotation.Attributes["filename"] = varFileName;
                        _service.Create(Annotation);

                        // Add the second report if relevant
                        if (sbErrorList.Length > 0)
                        {
                            byte[] filename2 = Encoding.ASCII.GetBytes(sbErrorList.ToString());
                            string encodedData2 = System.Convert.ToBase64String(filename2);
                            Entity Annotation2 = new Entity("annotation");
                            Annotation2.Attributes["objectid"] = new EntityReference("new_palmgoprap", varPrapID);
                            Annotation2.Attributes["objecttypecode"] = "new_palmgoprap";
                            Annotation2.Attributes["subject"] = "PRAP Extract";
                            Annotation2.Attributes["documentbody"] = encodedData2;
                            Annotation2.Attributes["mimetype"] = @"text / plain";
                            Annotation2.Attributes["notetext"] = "PRAP errors and warnings for " + cleanDate(varStartDatePr) + " to " + cleanDate(varEndDatePr);
                            Annotation2.Attributes["filename"] = varFileName2;
                            _service.Create(Annotation2);
                        }

                        //varTest += cleanDate(varStartDatePr) + " " + cleanDate(varEndDatePr);

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
    }
}

