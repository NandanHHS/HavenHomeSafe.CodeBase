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
    public class Budgets : IPlugin
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

                // Verify that the target entity represents the Palm SU Brokerage entity
                // If not, this plug-in was not registered correctly.
                // if (entity.LogicalName != "new_palmsubrokerage")
                //    return;

                try
                {
                    // Fixes:
                    // Remove ampersand from comments
                    // Compare values to Lower

                    EntityReference getEntity; // Object for entity references

                    string varTest = ""; // For debug
                    string dbBrokerage = ""; // brokerage description

                    // Only do this if the entity is the Palm SU Brokerage entity
                    if (entity.LogicalName == "new_palmsubrokerage")
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

                        // Fetch statements for database
                        // Get the information for all budgets
                        string dbBudgetList = @"
                            <fetch version='1.0' mapping='logical' distinct='false'>
                                <entity name='new_palmsubudget'>
                                <attribute name='new_brokerage'/>
                                <attribute name='new_startdate'/>
                                <attribute name='new_enddate'/>
                                <attribute name='new_palmsubudgetid'/>
                                </entity>
                            </fetch> ";

                        string varSpend = ""; // Spend amount
                        double varSpendDbl = 0; // Spend as double
                        string dbPalmSuBudgetId = ""; // Budget id
                        string dbStartDate = ""; // Budget start date
                        string dbEndDate = ""; // Budget end date
                        string dbBrokerageList = ""; // List of brokerages
                        EntityCollection result2; // Used for collection results processed below

                        // Get the fetch XML data and place in entity collection objects
                        EntityCollection result3 = _service.RetrieveMultiple(new FetchExpression(dbBudgetList));

                        // Loop through brokerages
                        foreach (var e in result3.Entities)
                        {
                            // Process the data as follows:
                            // If there is a formatted value for the field, use it
                            // Otherwise if there is a literal value for the field, use it
                            // Otherwise the value wasn't returned so set as nothing
                            if (e.Attributes.Contains("new_brokerage"))
                            {
                                getEntity = (EntityReference)e.Attributes["new_brokerage"];
                                dbBrokerage = getEntity.Id.ToString();
                            }
                            else if (e.FormattedValues.Contains("new_brokerage"))
                                dbBrokerage = e.FormattedValues["new_brokerage"];
                            else
                                dbBrokerage = "";

                            if (e.FormattedValues.Contains("new_startdate"))
                                dbStartDate = e.FormattedValues["new_startdate"];
                            else if (e.Attributes.Contains("new_startdate"))
                                dbStartDate = e.Attributes["new_startdate"].ToString();
                            else
                                dbStartDate = "";

                            if (e.FormattedValues.Contains("new_enddate"))
                                dbEndDate = e.FormattedValues["new_enddate"];
                            else if (e.Attributes.Contains("new_enddate"))
                                dbEndDate = e.Attributes["new_enddate"].ToString();
                            else
                                dbEndDate = "";

                            if (e.FormattedValues.Contains("new_palmsubudgetid"))
                                dbPalmSuBudgetId = e.FormattedValues["new_palmsubudgetid"];
                            else if (e.Attributes.Contains("new_palmsubudgetid"))
                                dbPalmSuBudgetId = e.Attributes["new_palmsubudgetid"].ToString();
                            else
                                dbPalmSuBudgetId = "";

                            // Convert date from American format to Australian format
                            dbStartDate = cleanDateAM(dbStartDate);
                            dbEndDate = cleanDateAM(dbEndDate);

                            //varTest += "Brokerage: " + dbBrokerage + "\r\n";
                            //varTest += "Budget: " + dbPalmSuBudgetId + "\r\n";
                            //varTest += "Dates: " + dbStartDate + " " + dbEndDate + " " + "\r\n";

                            // Get the total spend for the budget being processed
                            dbBrokerageList = @"
                                <fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                    <entity name='new_palmclientfinancial'>
                                    <attribute name='new_amount' alias='totalamount_sum' aggregate='sum'/> 
                                        <filter type='and'>
                                            <condition entityname='new_palmclientfinancial' attribute='new_brokerage' operator='eq' value='" + dbBrokerage + @"' />
                                            <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='le' value='" + dbEndDate + @"' />
                                            <condition entityname='new_palmclientfinancial' attribute='new_entrydate' operator='ge' value='" + dbStartDate + @"' />
                                            <filter type='or'>
                                                <condition entityname='new_palmclientfinancial' attribute='new_paidby' operator='ne' value='100000003' />
                                                <condition entityname='new_palmclientfinancial' attribute='new_voucherid' operator='like' value='DDPRAP#%' />
                                                <condition entityname='new_palmclientfinancial' attribute='new_voucherid' operator='like' value='DD#%' />
                                            </filter>
                                        </filter>
                                    </entity>
                                </fetch> ";

                            // Get the fetch XML data and place in entity collection objects
                            result2 = _service.RetrieveMultiple(new FetchExpression(dbBrokerageList));

                            // Reset variable
                            varSpend = "";

                            // Loop through spend records
                            foreach (var d in result2.Entities)
                            {
                                // get spend and clean formatting
                                if (d.FormattedValues.Contains("totalamount_sum"))
                                    varSpend = d.FormattedValues["totalamount_sum"];
                                else if (d.Attributes.Contains("totalamount_sum"))
                                    varSpend = d.Attributes["totalamount_sum"].ToString();
                                else
                                    varSpend = "";

                                varSpend = cleanString(varSpend, "double");
                            }

                            //varTest += "Spend: " + varSpend + "\r\n";

                            // Convert spend to double
                            if (!Double.TryParse(varSpend, out varSpendDbl))
                                varSpendDbl = 0;

                            // Update brokerage spent field for this brokerage
                            if (String.IsNullOrEmpty(dbPalmSuBudgetId) == false)
                            {
                                //varTest += "in\r\n";
                                Guid newGuid = Guid.Parse(dbPalmSuBudgetId);
                                Entity entBrok = new Entity("new_palmsubudget");
                                entBrok["new_palmsubudgetid"] = newGuid;
                                entBrok["new_brokspent"] = varSpendDbl;
                                _service.Update(entBrok);
                            }
                        }
                    }

                    //throw new InvalidPluginExecutionException("This plugin is working. Brokerage is " + dbBrokerage + "\r\n" + varTest);


                }
                //<snippetFollowupPlugin3>
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in the FollowupPlugin plug-in.", ex);
                }
                //</snippetFollowupPlugin3>

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

        // Fix American date issue
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
    }
}


