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
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Client;
using System.Net.Http;
using System.Net;
using System.Text;
using Newtonsoft.Json;

namespace BaseProfileCopy
{
    public static class CopyBaseProfile
    {
        [FunctionName("CopyBaseProfile")]
        public static async Task<HttpResponseMessage> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            
            //Get the profile id that we need to copy
            string id = req.Query["id"];

            //Client id, client secret , environment
            const string clientId = "a9545b29-056e-4eab-8f31-911e9f24c52a";
            const string clientSecret = "HhXKcdJuY6Kid4EFdLhFBXmeIFHw2SuC";
            const string environment = "https://operations-metcash-pp-price.crm6.dynamics.com";
            var connectionString = @$"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";
            using var serviceClient = new ServiceClient(connectionString);
            var idCreated = "";
            
            try
            {

                //Retrieve record that needs to be copied
                Entity baseprofile = new Entity("met_pricingprofile");

                Entity retrievedAccount = serviceClient.Retrieve(
                   entityName: baseprofile.LogicalName,
                   id: new Guid(id),
                   columnSet: new ColumnSet(true)
               );
                Entity baseprofileCreate = new Entity("met_pricingprofile");
                
                //Create entry for pricing profile 

                baseprofileCreate["met_baseformarkup"] = retrievedAccount.GetAttributeValue("met_baseformarkup");
                baseprofileCreate["met_baseprofileid"] =  retrievedAccount.GetAttributeValue("met_baseprofileid") != null ? new EntityReference("met_pricingprofile", ((EntityReference)retrievedAccount.GetAttributeValue("met_baseprofileid")).Id):null;
                baseprofileCreate["met_profiletypeid"] = new EntityReference("met_pricingprofiletype", ((EntityReference)retrievedAccount.GetAttributeValue("met_profiletypeid")).Id);
                baseprofileCreate["met_customerid"] = retrievedAccount.GetAttributeValue("met_customerid")!=null? new EntityReference("account", ((EntityReference)retrievedAccount.GetAttributeValue("met_customerid")).Id):null;
                baseprofileCreate["met_siteid"] = retrievedAccount.GetAttributeValue("met_siteid") !=null ?new EntityReference("site", ((EntityReference)retrievedAccount.GetAttributeValue("met_siteid")).Id):null;
                baseprofileCreate["met_divisionid"] = retrievedAccount.GetAttributeValue("met_divisionid") != null ? new EntityReference("cdm_company", ((EntityReference)retrievedAccount.GetAttributeValue("met_divisionid")).Id):null;
                baseprofileCreate["met_fixedpriceprofile"] = retrievedAccount.GetAttributeValue("met_fixedpriceprofile");
                baseprofileCreate["met_costbase"] = retrievedAccount.GetAttributeValue("met_costbase");
                baseprofileCreate["met_changedate"] = retrievedAccount.GetAttributeValue("met_changedate");
                baseprofileCreate["met_markuptype"] = retrievedAccount.GetAttributeValue("met_markuptype");
                baseprofileCreate["met_baseformarkup"] = retrievedAccount.GetAttributeValue("met_baseformarkup");
                baseprofileCreate["met_shortdescription"] = retrievedAccount.GetAttributeValue("met_shortdescription");
                baseprofileCreate["met_profiledescription"] = retrievedAccount.GetAttributeValue("met_profiledescription");
                baseprofileCreate["met_profilenumber"] = retrievedAccount.GetAttributeValue("met_profilenumber");
                baseprofileCreate["met_hostindicator"] = retrievedAccount.GetAttributeValue("met_hostindicator");
                baseprofileCreate["met_fixedpriceprofile"] = retrievedAccount.GetAttributeValue("met_fixedpriceprofile");
                baseprofileCreate["met_noparticipation"] = retrievedAccount.GetAttributeValue("met_noparticipation");
                baseprofileCreate["met_iscopy"] = true;
                baseprofileCreate["met_validfrom"] = DateTime.Now;
                baseprofileCreate["met_originalprofileid"] = new EntityReference("met_pricingprofile", new Guid(id));
                baseprofileCreate["statuscode"] = new OptionSetValue(862200000);
                
                //Create the pricing profile so as to associate
                var baseprofileGuid = serviceClient.Create(baseprofileCreate);
                idCreated = baseprofileGuid.ToString();

                //Filter on the basis of pricing profile and active
                ConditionExpression condPricingProfileID = new ConditionExpression
                {
                    AttributeName = "met_pricingprofileid",
                    Operator = ConditionOperator.Equal
                };
                condPricingProfileID.Values.Add(id);

                ConditionExpression conditionForActive = new ConditionExpression() { AttributeName = "statecode", Operator = ConditionOperator.Equal }; ;
                conditionForActive.Values.Add(0);

                FilterExpression filterCondition = new FilterExpression();
                filterCondition.Conditions.Add(condPricingProfileID);
                filterCondition.Conditions.Add(conditionForActive);

                QueryExpression query = new QueryExpression("met_pricingprofileline");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria.AddFilter(filterCondition);
                
                EntityCollection retrieveMultipleEntity = serviceClient.RetrieveMultiple(query);


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
                
                    //Loop through each line and create pricing profile line
                    foreach (var pricinglines in retrieveMultipleEntity.Entities)
                {


                    Entity baseprofilelineCreate = new Entity("met_pricingprofileline");
                    baseprofilelineCreate["met_baseprofileid"] = pricinglines.Attributes.ContainsKey("met_baseprofileid") && pricinglines.Attributes["met_baseprofileid"] != null ? new EntityReference("met_pricingprofile", ((EntityReference)pricinglines.Attributes["met_baseprofileid"]).Id) : null;
                    baseprofilelineCreate["met_productid"] = pricinglines.Attributes.ContainsKey("met_productid") && pricinglines.Attributes["met_productid"] != null ? new EntityReference("product", ((EntityReference)pricinglines.Attributes["met_productid"]).Id) : null;
                    baseprofilelineCreate["met_productcategoryid"] = pricinglines.Attributes.ContainsKey("met_productcategoryid") && pricinglines.Attributes["met_productcategoryid"] != null ? new EntityReference("msdyn_productcategory", ((EntityReference)pricinglines.Attributes["met_productcategoryid"]).Id) : null;
                    baseprofilelineCreate["met_attachmenttypeid"] = pricinglines.Attributes.ContainsKey("met_attachmenttypeid") && pricinglines.Attributes["met_attachmenttypeid"] != null ? new EntityReference("met_attachmenttype", ((EntityReference)pricinglines.Attributes["met_attachmenttypeid"]).Id) : null;
                    baseprofilelineCreate["met_pricingprofileid"] = new EntityReference("met_pricingprofile", (baseprofileGuid));
                    baseprofilelineCreate["met_iscopy"] = true;
                    baseprofilelineCreate["met_profilelinename"] = pricinglines.Attributes.ContainsKey("met_profilelinename") ? pricinglines.Attributes["met_profilelinename"] : null;
                    baseprofilelineCreate["met_markuptype"] = pricinglines.Attributes.ContainsKey("met_markuptype")? pricinglines.Attributes["met_markuptype"]:null;
                    baseprofilelineCreate["met_earlierpercentage"] = pricinglines.Attributes.ContainsKey("met_earlierpercentage") ? pricinglines.Attributes["met_earlierpercentage"]:null; ;
                    baseprofilelineCreate["exchangerate"] = pricinglines.Attributes.ContainsKey("exchangerate") ? pricinglines.Attributes["exchangerate"]:null; 
                    baseprofilelineCreate["met_percentage"] = pricinglines.Attributes.ContainsKey("met_percentage") ? pricinglines.Attributes["met_percentage"]:null; 
                    baseprofilelineCreate["met_wholesaleprice"] = pricinglines.Attributes.ContainsKey("met_wholesaleprice") ? pricinglines.Attributes["met_wholesaleprice"]:null;
                    baseprofilelineCreate["met_srpprice"] = pricinglines.Attributes.ContainsKey("met_srpprice") ? pricinglines.Attributes["met_srpprice"]:null;
                    baseprofilelineCreate["met_earliersrpprice"] = pricinglines.Attributes.ContainsKey("met_earliersrpprice") ? pricinglines.Attributes["met_earliersrpprice"]:null;
                    baseprofilelineCreate["met_originalprofilelineid"] = new EntityReference("met_pricingprofile", new Guid(pricinglines.Attributes["met_pricingprofilelineid"].ToString()));

                  


                    var createRequest = new CreateRequest()
                    {
                        Target = baseprofilelineCreate
                    };
                    if (request.Requests.Count >= 997)
                    {
                        BulkExecute(serviceClient, request);
                        request = BulkExecuteRequest();
                    }
                    else
                    {
                        request.Requests.Add(createRequest);
                    }
                }

                //Execute Multiple request
                var response = (ExecuteMultipleResponse)serviceClient.Execute(request);

                foreach (var r in response.Responses)
                {
                    if (r.Response != null)
                        Console.WriteLine("Success********" + r.Response);
                    else if (r.Fault != null)
                        Console.WriteLine(r.Fault);
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }
            var myObj = new { baseid = idCreated};
            var jsonToReturn = JsonConvert.SerializeObject(myObj);
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(jsonToReturn, Encoding.UTF8, "application/json")
            };
        }

        public static ExecuteMultipleRequest BulkDeleteRequest()
        {
            var multipleRequest = new ExecuteMultipleRequest()
            {
                // Assign settings that define execution behavior: continue on error, return responses.
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                // Create an empty organization request collection.
                Requests = new OrganizationRequestCollection()

                
            };
            return multipleRequest;
            // Execute all the requests in the request collection using a single web method call.
        }

        public static ExecuteMultipleRequest BulkExecuteRequest()
        {
            var multipleRequest = new ExecuteMultipleRequest()
            {
                // Assign settings that define execution behavior: continue on error, return responses.
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                // Create an empty organization request collection.
                Requests = new OrganizationRequestCollection()


            };
            return multipleRequest;
            // Execute all the requests in the request collection using a single web method call.
        }

        public static void BulkExecute(ServiceClient service, ExecuteMultipleRequest multipleRequest)
        {
            ExecuteMultipleResponse multipleResponse = (ExecuteMultipleResponse)service.Execute(multipleRequest);
        }
    }
}
