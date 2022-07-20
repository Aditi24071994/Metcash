using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;



namespace FunctionApp11
{
    public class Sample1
    {
        public string DataFromDataverse { get; set; }


    }
    public static class Function2
    {
        [FunctionName("Function2")]

        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.System, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP tr3igger function processed a request.");
            string id = req.Query["id"];
            /*-----------------------------------------
             * The next code will get the client id , client secret and environment url and connect to table testFY
             * testFy - crb3f_testfy , column considered - crb3f_datecustom
             * taking the top 10 count 
             * convert to csv 
             * save in local for reference
             */
            const string clientId = "a9545b29-056e-4eab-8f31-911e9f24c52a";
            const string clientSecret = "HhXKcdJuY6Kid4EFdLhFBXmeIFHw2SuC";
            const string environment = "https://aditigautamsenvironment.crm.dynamics.com";
            var connectionString = @$"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";
            using var serviceClient = new ServiceClient(connectionString);
            Entity testfy = new Entity("crb3f_testfy");
            testfy.Id = new Guid(id);
            testfy["crb3f_name"] = "updated";
            serviceClient.Update(testfy);
            log.LogInformation("End");
            return null;
        }
    }
}
