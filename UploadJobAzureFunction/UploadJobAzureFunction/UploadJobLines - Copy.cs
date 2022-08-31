using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using System.Linq;
using Microsoft.VisualBasic.FileIO;
using System.Data;
using Microsoft.Xrm.Sdk.Messages;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace UploadJobAzureFunction
{
    public static class UploadJobLines11
    {
        [FunctionName("UploadJobLines1")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest headerRequest,
            ILogger log)
        {
            log.LogInformation("Function triggered");

            string id = headerRequest.Query["id"];
            string type = headerRequest.Query["type"];
            int countOfCreated = 0, countOfUpdated = 0;

            var clientId = "a9545b29-056e-4eab-8f31-911e9f24c52a";
            const string clientSecret = "HhXKcdJuY6Kid4EFdLhFBXmeIFHw2SuC";
            const string environment = "https://operations-metcash-pp-price.crm6.dynamics.com";
            var connectionString = @$"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";

            var serviceClient = new ServiceClient(connectionString);
            var relatedEntities = new EntityReferenceCollection();
            var requestForMultipleFetch = new ExecuteMultipleRequest()
            {
                Requests = new OrganizationRequestCollection(),
                Settings = new ExecuteMultipleSettings
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                }
            };
            
            var jobLine = new Entity("product");
            jobLine["msdyn_productnumber"] = "4446068";
            jobLine["name"] = "testrecord";
            jobLine["defaultuomid"] = new EntityReference("uom", Guid.Parse("4b922227-940c-ed11-b83d-00224810bcf8"));
            jobLine["defaultuomid"] = new EntityReference("uom", Guid.Parse("4b922227-940c-ed11-b83d-00224810bcf8"));


            var jobLineCreate = new CreateRequest()
            {
                Target = jobLine
            };
            requestForMultipleFetch.Requests.Add(jobLineCreate);
            var response = (ExecuteMultipleResponse)serviceClient.Execute(requestForMultipleFetch);
            Console.WriteLine("After execute");

            foreach (var r in response.Responses)
            {
                if (r.Response != null)


                    Console.WriteLine("Success" + r.Response + "----" + r.GetType().Name);
                else if (r.Fault != null)
                    Console.WriteLine(r.Fault);
            }
            return null;
        }
    }
}