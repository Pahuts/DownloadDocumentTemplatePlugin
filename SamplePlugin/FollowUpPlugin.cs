using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using System.Net.Http.Headers;
using Newtonsoft.Json; // for Json Functions
using System.ServiceModel.Description;
using System.Net;
using Microsoft.Xrm.Sdk.Client;



namespace SamplePlugin
{
    public class CustomAction : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            // Create a new record
            Entity contact = new Entity("contact");
            contact["firstname"] = "Bob";
            contact["lastname"] = "Smith";
            Guid contactId = service.Create(contact);
        }
    }

    public class DownloadDocumentTemplate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            string documentTemplateId = context.InputParameters["TemplateId"].ToString();
            string recordId = context.InputParameters["RecordId"].ToString();
            string opptyRecordId = context.InputParameters["OpptyRecordId"].ToString();
            int typeCode = Convert.ToInt32(context.InputParameters["TypeCode"]);

            CreatePDFNote(service, documentTemplateId, recordId, opptyRecordId, typeCode);
        }

        public static void CreatePDFNote(IOrganizationService _service, string documentTemplateId, string recordId, string opptyRecordId, int entityTypeCode)
        {
            // Create new Organization service with admin user to call "ExportPdfDocument" message
            IOrganizationService pdfService = GetExportPDFService();

            try
            {
                OrganizationRequest request = new OrganizationRequest("ExportPdfDocument");
                request["EntityTypeCode"] = entityTypeCode;
                request["SelectedTemplate"] = new EntityReference("documenttemplate", new Guid(documentTemplateId));
                List<Guid> records = new List<Guid> { new Guid(recordId) };
                request["SelectedRecords"] = JsonConvert.SerializeObject(records);

                OrganizationResponse pdfResponse = (OrganizationResponse)pdfService.Execute(request);

                //Write to file
                string b64File = Convert.ToBase64String((byte[])pdfResponse["PdfFile"]);

                // Create note by using the above base 64 string / create email attachment and send it to customer .

                Entity Annotation = new Entity("annotation");
                Annotation.Attributes["subject"] = "Fetched Contact Details";
                Annotation.Attributes["documentbody"] = b64File;
                Annotation.Attributes["objectid"] = new EntityReference("opportunity", new Guid(opptyRecordId));
                Annotation.Attributes["mimetype"] = @"application/pdf";
                Annotation.Attributes["notetext"] = "Download Contact Details";
                Annotation.Attributes["filename"] = "ContactDetails.pdf";
                _service.Create(Annotation);
            }
            catch (Exception ex)
            {
            }
        }

        public static IOrganizationService GetExportPDFService()
        {
            IOrganizationService _service = null;
            if (_service == null)
            {
                string serviceURL = "https://aimdev.crm5.dynamics.com/XRMServices/2011/Organization.svc";
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = "crm.sa@aim.edu";
                credentials.UserName.Password = ";q4w:$RJK'4S";
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                Uri serviceuri = new Uri(serviceURL);

                try
                {
                    OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceuri, null, credentials, null);
                    proxy.EnableProxyTypes();
                    _service = (IOrganizationService)proxy;

                }
                catch (Exception ex)
                {

                }
            }
            return _service;
        }
    }

    public class ValidateAccountName : IPlugin
    {
        //Invalid names from unsecure configuration
        private List<string> invalidNames = new List<string>();

        // Constructor to capture the unsecure configuration
        public ValidateAccountName(string unsecure)
        {
            // Parse the configuration data and set invalidNames
            if (!string.IsNullOrWhiteSpace(unsecure))
                unsecure.Split(',').ToList().ForEach(s =>
                {
                    invalidNames.Add(s.Trim());
                });
        }
        public void Execute(IServiceProvider serviceProvider)
        {

            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            try
            {

                // Obtain the execution context from the service provider.  
                IPluginExecutionContext context = (IPluginExecutionContext)
                    serviceProvider.GetService(typeof(IPluginExecutionContext));

                // Verify all the requirements for the step registration
                if (context.InputParameters.Contains("Target") && //Is a message with Target
                    context.InputParameters["Target"] is Entity && //Target is an entity
                    ((Entity)context.InputParameters["Target"]).LogicalName.Equals("account") && //Target is an account
                    ((Entity)context.InputParameters["Target"])["name"] != null && //account name is passed
                    context.MessageName.Equals("Update") && //Message is Update
                    context.PreEntityImages["a"] != null && //PreEntityImage with alias 'a' included with step
                    context.PreEntityImages["a"]["name"] != null) //account name included with PreEntityImage with step
                {
                    // Obtain the target entity from the input parameters.  
                    var entity = (Entity)context.InputParameters["Target"];
                    var newAccountName = (string)entity["name"];
                    var oldAccountName = (string)context.PreEntityImages["a"]["name"];

                    if (invalidNames.Count > 0)
                    {
                        tracingService.Trace("ValidateAccountName: Testing for {0} invalid names:", invalidNames.Count);

                        if (invalidNames.Contains(newAccountName.ToLower().Trim()))
                        {
                            tracingService.Trace("ValidateAccountName: new name '{0}' found in invalid names.", newAccountName);

                            // Test whether the old name contained the new name
                            if (!oldAccountName.ToLower().Contains(newAccountName.ToLower().Trim()))
                            {
                                tracingService.Trace("ValidateAccountName: new name '{0}' not found in '{1}'.", newAccountName, oldAccountName);

                                string message = string.Format("You can't change the name of this account from '{0}' to '{1}'.", oldAccountName, newAccountName);

                                throw new InvalidPluginExecutionException(message);
                            }

                            tracingService.Trace("ValidateAccountName: new name '{0}' found in old name '{1}'.", newAccountName, oldAccountName);
                        }

                        tracingService.Trace("ValidateAccountName: new name '{0}' not found in invalidNames.", newAccountName);
                    }
                    else
                    {
                        tracingService.Trace("ValidateAccountName: No invalid names passed in configuration.");
                    }
                }
                else
                {
                    tracingService.Trace("ValidateAccountName: The step for this plug-in is not configured correctly.");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("SamplePlugin: {0}", ex.ToString());
                throw;
            }
        }
    }
public class FollowupPlugin : IPlugin
	{
		public void Execute(IServiceProvider serviceProvider)
		{
			// Obtain the tracing service
			ITracingService tracingService =
			(ITracingService)serviceProvider.GetService(typeof(ITracingService));

			// Obtain the execution context from the service provider.  
			IPluginExecutionContext context = (IPluginExecutionContext)
				serviceProvider.GetService(typeof(IPluginExecutionContext));

			// The InputParameters collection contains all the data passed in the message request.  
			if (context.InputParameters.Contains("Target") &&
				context.InputParameters["Target"] is Entity)
			{
				// Obtain the target entity from the input parameters.  
				Entity entity = (Entity)context.InputParameters["Target"];

				// Obtain the organization service reference which you will need for  
				// web service calls.  
				IOrganizationServiceFactory serviceFactory =
					(IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
				IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

				try
				{
					// Plug-in business logic goes here.
					// Create a task activity to follow up with the account customer in 7 days.
					Entity followup = new Entity("task");
					followup["subject"] = "Send e-mail to the new customer.";
					followup["description"] = "Follow up with the customer. Check if there are any new issues that need resolution";
					followup["scheduledstart"] = DateTime.Now.AddDays(7);
					followup["scheduledend"] = DateTime.Now.AddDays(7);
					followup["category"] = context.PrimaryEntityName; 

					// Refer to the account in the task activity.
					if (context.OutputParameters.Contains("id"))
					{
						Guid regardingobjectid = new Guid(context.OutputParameters["id"].ToString());
						string regardingobjectidType = "account";

						followup["regardingobjectid"] = new EntityReference(regardingobjectidType, regardingobjectid);
					}

					// Create the task in Microsoft Dynamics CRM.
					tracingService.Trace("FollowupPlugin: Creating the task activity.");
					service.Create(followup);
				}

				catch (FaultException<OrganizationServiceFault> ex)
				{
					throw new InvalidPluginExecutionException("An error occurred in FollowUpPlugin.", ex);
				}

				catch (Exception ex)
				{
					tracingService.Trace("FollowUpPlugin: {0}", ex.ToString());
					throw;
				}
			}
		}
	}
}
