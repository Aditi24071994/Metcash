using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Xrm.Sdk;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk.Messages;

namespace UpdateCustomerProfileStatus
{
    /*
     * 
     * 
     * */
    public static class UpdateCustomerProfile
    {
        [FunctionName("UpdateCustomerProfile")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("UpdateCustomerProfile HTTP trigger function processed a request.");

            string recordId = req.Query["id"];
            string ownerID = req.Query["owner"];
            log.LogInformation(ownerID);
            log.LogInformation("  {  \"actions\": [   {                \"title\": \"Customer Profile updated\",        \"data\": {                    \"url\": \"?pagetype=entityrecord&etn=met_pricingprofile&id=" + recordId + "\",            \"navigationTarget\": \"inline\"               }            }      ]     }");

            const string clientId = "a9545b29-056e-4eab-8f31-911e9f24c52a";
            const string clientSecret = "HhXKcdJuY6Kid4EFdLhFBXmeIFHw2SuC";
            const string environment = "https://operations-metcash-pp-price.crm6.dynamics.com";
            var connectionString = @$"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";
            var serviceClient = new ServiceClient(connectionString);

            //Update the isready flag
            var updatedProfileline = new Entity("met_pricingprofile",new Guid(recordId));
            updatedProfileline["met_isready"] = true;
            serviceClient.Update(updatedProfileline);

            //Send In app notification
            var inappNotification = new Entity("appnotification");

            inappNotification["title"] = "Customer Profile updated";
            inappNotification["ownerid"] = new EntityReference("systemuser", new Guid(ownerID));
            inappNotification["icontype"] = new OptionSetValue(100000000);
            inappNotification["toasttype"] = new OptionSetValue(200000000); 
            inappNotification["data"] = "  {  \"actions\": [   {                \"title\": \"Customer Profile updated\",        \"data\": {                    \"url\": \"?pagetype=entityrecord&etn=met_pricingprofile&id=" + recordId + "\",            \"navigationTarget\": \"inline\"               }            }      ]     }";
            

            Guid appNotificationId = serviceClient.Create(inappNotification);


            return null;
        }
    }
}
