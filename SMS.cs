// Plugin namespaces
using System;
using System.ServiceModel;

// Microsoft Dynamics CRM namespace(s)
using Microsoft.Xrm.Sdk;
using System.Net;
using System.Web;
using System.Security;
using System.Net.Mail;
using System.Text;
using System.IO;

[assembly: AllowPartiallyTrustedCallers]
// Sample plugin structure adapted but all code modified
namespace Microsoft.Crm.Sdk.Samples
{
    public class SendSMS : IPlugin
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

                // Verify that the target entity represents the Palm Client SMS entity
                // If not, this plug-in was not registered correctly.
                // if (entity.LogicalName != "new_palmclientsms")
                //    return;

                try
                {
                    string varTest = ""; // Message from form
                    bool varT2 = false; // Send SMS checkbox

                    // Only do this if the entity is the Palm Client SMS entity
                    if (entity.LogicalName == "new_palmclientsms")
                    {
                        varTest = entity.GetAttributeValue<string>("new_message");
                        varT2 = entity.GetAttributeValue<bool>("new_sendsms");
                    }

                    // Do this if the send SMS checkbox is ticked
                    if (varT2 == true)
                    {
                        // Get hardcoded SMS credentials
                        string varUsername = "trevor.perri@haven.org.au";
                        string varPassword = WebUtility.UrlEncode("Nic945#");

                        string varSource = "61438717021";
                        string varPhone = "0409131107";

                        // Get message and limit to 255 characters
                        if (String.IsNullOrEmpty(varTest) == true)
                            varTest = "empty";
                        if (varTest.Length > 255)
                            varTest = varTest.Substring(0, 255);

                        // Encode SMS message
                        string varMessage = WebUtility.UrlEncode(varTest);

                        // Create post data
                        string api = "username=" + varUsername + "&password=" + varPassword + "&source=" + varSource + "&destination=" + varPhone + "&text=" + varMessage + "";

                        // Create HTTP request and send SMS
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://tim.telstra.com/cgphttp/servlet/sendmsg");
                        request.Method = "POST";
                        request.ContentType = "application/x-www-form-urlencoded";

                        byte[] bytedata = Encoding.UTF8.GetBytes(api);
                        request.ContentLength = bytedata.Length;

                        Stream dataStream = request.GetRequestStream();
                        dataStream.Write(bytedata, 0, bytedata.Length);
                        dataStream.Close();

                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
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
    }
}


