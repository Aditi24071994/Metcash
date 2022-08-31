using System;
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

namespace UploadJobAzureFunction
{
    public static class MapJobLines
    {
        [FunctionName("MapJobLines")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest headerRequest,
            ILogger log)
        {
            log.LogInformation("MAp Job lines Function triggered");

            string id = headerRequest.Query["id"];
            string type = headerRequest.Query["type"];
            int countOfCreated = 0, countOfUpdated = 0;

            var clientId = "a9545b29-056e-4eab-8f31-911e9f24c52a";
            const string clientSecret = "HhXKcdJuY6Kid4EFdLhFBXmeIFHw2SuC";
            const string environment = "https://operations-metcash-pp-price.crm6.dynamics.com";
            var connectionString = @$"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";

            var serviceClient = new ServiceClient(connectionString);
            var fetchxmlForJobLine = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>  <entity name='met_uploadjobprofileline'><attribute name='met_uploadjobprofilelineid' /><attribute name='met_uploadjobprofilename' /><attribute name='met_pricingprofileid' /><attribute name='met_newsrp' /><attribute name='met_itemid' /><order attribute='met_uploadjobprofilename' descending='false' /><filter type='and'>  <condition attribute='met_uploadjobid' operator='eq' uitype='met_uploadjob' value='{" + id+"}' />  <condition attribute='statuscode' operator='eq' value='862200000' /></filter>  </entity></fetch>";

            EntityCollection colMasterJobLines = PopulateMastersForPRofileLines("met_pricingprofileline", serviceClient, fetchxmlForJobLine);

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
           
            foreach (var cols in colMasterJobLines.Entities)
            {
                Entity profilelineAction = new();
                KeyAttributeCollection altKey = new KeyAttributeCollection();
                Console.WriteLine(cols.Attributes["met_newsrp"].ToString());
                altKey.Add("met_productid", new EntityReference("product", ((EntityReference)cols.Attributes["met_itemid"]).Id));
                altKey.Add("met_pricingprofileid", new EntityReference("met_pricingprofile", ((EntityReference)cols.Attributes["met_pricingprofileid"]).Id));
                Console.WriteLine("Beforse execute");

                profilelineAction = new Entity("met_pricingprofileline", altKey);
                profilelineAction["met_srpprice"] = cols.Attributes["met_newsrp"];
                Console.WriteLine("Beforse execute");

                UpsertRequest upsertRequest = new()
                {
                    Target = profilelineAction
                };
                requestForMultipleFetch.Requests.Add(upsertRequest);
            }


            Console.WriteLine("Before execute");
            var response = (ExecuteMultipleResponse)serviceClient.Execute(requestForMultipleFetch);
            Console.WriteLine("After execute");


            foreach (var r in response.Responses)
            {
                if (r.Response != null)
                {
                    if(r.Response.ToString() == "Microsoft.Xrm.Sdk.Messages.UpsertResponse")
                    {
                        if (((UpsertResponse)r.Response).RecordCreated)
                        {
                            countOfCreated++;
                        }
                        else
                        {
                            countOfUpdated++;
                        }
                    }
                    
                }
                else if (r.Fault != null)
                {
                    Console.WriteLine(r.Fault);
                }
            }

            var jobUpload = new Entity("met_uploadjob", new Guid(id));
           
            jobUpload["statuscode"] = new OptionSetValue(862200002);
            jobUpload["met_recordscreated"] = countOfCreated;
            jobUpload["met_recordsupdated"] = countOfUpdated;
            var jobUploadCreate = new UpdateRequest()
            {
                Target = jobUpload
            };


            serviceClient.Update(jobUpload);
            Console.WriteLine("Created" + countOfCreated.ToString());
            Console.WriteLine("Updated" + countOfUpdated.ToString());

            return null;
        }
        public static EntityCollection PopulateMastersForPRofileLines(string name, ServiceClient service, String fetchString)
        {
            EntityCollection tempCol = new();
            EntityCollection MainCol = new();

            var pageNumber = 1;
            var pagingCookie = string.Empty;
            ConditionExpression conditionForActive = new()
            {
                AttributeName = "statecode",
                Operator = ConditionOperator.Equal
            };
            conditionForActive.Values.Add(2);
            FilterExpression filterForActive = new();
            filterForActive.Conditions.Add(conditionForActive);

            if (!string.IsNullOrEmpty(fetchString))
            {
                Console.WriteLine("in if " + name);

                QueryExpression query = new(name)
                {
                    ColumnSet = new ColumnSet(true)
                };
                do
                {
                    if (pageNumber != 1)
                    {
                        query.PageInfo.PageNumber = pageNumber;
                        query.PageInfo.PagingCookie = pagingCookie;
                    }
                    tempCol = service.RetrieveMultiple(new FetchExpression(fetchString));
                    if (tempCol.MoreRecords)
                    {
                        pageNumber++;
                        pagingCookie = tempCol.PagingCookie;
                    }
                    MainCol.Entities.AddRange(tempCol.Entities);
                } while (tempCol.MoreRecords);
                //query.Criteria.AddFilter(filterForActive);
                MainCol = service.RetrieveMultiple(new FetchExpression(fetchString));
            }
            else
            {
                Console.WriteLine("in else " + name);
                QueryExpression query = new(name)
                {
                    ColumnSet = new ColumnSet(true)
                };
                do
                {
                    if (pageNumber != 1)
                    {
                        query.PageInfo.PageNumber = pageNumber;
                        query.PageInfo.PagingCookie = pagingCookie;
                    }
                    tempCol = service.RetrieveMultiple(query);
                    if (tempCol.MoreRecords)
                    {
                        pageNumber++;
                        pagingCookie = tempCol.PagingCookie;
                    }
                    MainCol.Entities.AddRange(tempCol.Entities);
                } while (tempCol.MoreRecords);
                //query.Criteria.AddFilter(filterForActive);
                MainCol = service.RetrieveMultiple(query);
            }
            return MainCol;
        }
    }
}