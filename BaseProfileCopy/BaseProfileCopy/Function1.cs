using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;

namespace BaseProfileCopy
{
    public static class Function1
    {
        [FunctionName("Functiontest")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string id = req.Query["id"];

            
            const string clientId = "a9545b29-056e-4eab-8f31-911e9f24c52a";
            const string clientSecret = "HhXKcdJuY6Kid4EFdLhFBXmeIFHw2SuC";
            const string environment = "https://operations-metcash-pp-price.crm6.dynamics.com";
            var connectionString = @$"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";
            using var serviceClient = new ServiceClient(connectionString);
            try
            {

                Entity baseprofile = new Entity("met_pricingprofile");

                Entity retrievedAccount = serviceClient.Retrieve(
                   entityName: baseprofile.LogicalName,
                   id: new Guid(id),
                   columnSet: new ColumnSet("met_profilenumber", "met_pricingprofileid")
               );
                Entity baseprofileCreate = new Entity("met_pricingprofile");
               
                
                baseprofileCreate["met_profilenumber"] = retrievedAccount["met_profilenumber"]+"  Updated";
                baseprofileCreate.Id =  serviceClient.Create(baseprofileCreate);
                var baseprofileGuid = baseprofileCreate.Id.ToString();

                ConditionExpression condition1 = new ConditionExpression();
                condition1.AttributeName = "met_pricingprofileid";
                condition1.Operator = ConditionOperator.Equal;
                condition1.Values.Add(id);

               
                FilterExpression filter1 = new FilterExpression();
                filter1.Conditions.Add(condition1);
                

                QueryExpression query = new QueryExpression("met_pricingprofileline");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria.AddFilter(filter1);
                EntityCollection result1 = serviceClient.RetrieveMultiple(query);
                var relatedEntities = new EntityReferenceCollection();
                var request = new ExecuteMultipleRequest()
                {
                    Requests = new OrganizationRequestCollection(),
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    }
                };

                foreach (var a in result1.Entities)
                {

                    Console.WriteLine("Nam: " + a.LogicalName);
                    foreach (var a1 in a.Attributes)
                    {
                        Console.WriteLine("Attr: " + a1.Key);
                        
                    }
                   
                }
                foreach (var a in result1.Entities)
                {

                    Console.WriteLine("met_pricingprofilelineid: " + a.Attributes["met_pricingprofilelineid"]);
                    Entity baseprofilelineCreate = new Entity("met_pricingprofileline");
                    baseprofilelineCreate["met_profilelinename"] = a.Attributes["met_profilelinename"];
                    baseprofilelineCreate["met_pricingprofileid"] = new EntityReference("met_pricingprofile", new Guid(baseprofileGuid));

                    
                    //baseprofilelineCreate["met_baseprofileid"] = new EntityReference("met_pricingprofile",((EntityReference)a.Attributes["met_baseprofileid"]).Id);
                    //baseprofilelineCreate["transactioncurrencyid"] = ((EntityReference)a.Attributes["transactioncurrencyid"]).Id;
                   // baseprofilelineCreate["exchangerate"] = a.Attributes["exchangerate"];
                    baseprofilelineCreate["met_iscopy"] = true;
                    //baseprofilelineCreate["met_productid"] = new EntityReference("product", ((EntityReference)a.Attributes["met_productid"]).Id);
                    //baseprofilelineCreate["met_productcategoryid"] = new EntityReference("msdyn_productcategory", ((EntityReference)a.Attributes["msdyn_productcategory"]).Id);
                    //baseprofilelineCreate["met_attachmenttypeid"] = new EntityReference("met_attachmenttype", ((EntityReference)a.Attributes["met_attachmenttype"]).Id);
                    // baseprofilelineCreate["met_markuptype"] = a.Attributes["met_markuptype"];
                    
                    baseprofilelineCreate["met_originalprofilelineid"] = new EntityReference("met_pricingprofileline", new Guid(a.Attributes["met_pricingprofilelineid"].ToString()));
                    











                    var createRequest = new CreateRequest()
                    {
                        Target = baseprofilelineCreate
                    };
                    request.Requests.Add(createRequest);
                    //baseprofilelineCreate["met_profilelinename"] = baseprofileCreate.Id;
                    
                }


                Console.WriteLine("Before execute");
                var response = (ExecuteMultipleResponse)serviceClient.Execute(request);
                Console.WriteLine("After execute");

                foreach (var r in response.Responses)
                {
                    if (r.Response != null)
                        Console.WriteLine("Success"+r.Response);
                    else if (r.Fault != null)
                        Console.WriteLine(r.Fault);
                }
                








            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }


            return null;
        }
    }
}
