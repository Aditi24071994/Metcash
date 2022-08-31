using System;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using System.Linq;

namespace SyncApprovedChanges
{
    internal class SyncApprovedChanges
    {
        static void Main()
        {
            
            const string clientId = "a9545b29-056e-4eab-8f31-911e9f24c52a";
            const string clientSecret = "HhXKcdJuY6Kid4EFdLhFBXmeIFHw2SuC";
            const string environment = "https://operations-metcash-pp-price.crm6.dynamics.com";
            var connectionString = $@"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";
            Console.WriteLine("connectionString: " + connectionString.ToString());
            var serviceClient = new ServiceClient(connectionString);
            try
            {
                /********************************
                 * Prepare filter conditions
                 * ** validtime<= currenttime
                 * ** active record
                 * ** is copy = Yes
                 * ** is approved
                 ********************************/
                ConditionExpression validFromConditionForBase = new ConditionExpression() { AttributeName = "met_validfrom", Operator = ConditionOperator.LessEqual };
                validFromConditionForBase.Values.Add(DateTime.Now);
                ConditionExpression isCopyConditionForBase = new ConditionExpression() { AttributeName = "met_iscopy", Operator= ConditionOperator.Equal };
                isCopyConditionForBase.Values.Add(true);
                ConditionExpression isActiveConditionForBase = new ConditionExpression() { AttributeName = "statecode", Operator = ConditionOperator.Equal }; ;
                isActiveConditionForBase.Values.Add(0);
                ConditionExpression isApprovedConditionForBase = new ConditionExpression() { AttributeName = "statuscode", Operator = ConditionOperator.Equal }; ;
                isApprovedConditionForBase.Values.Add(862200001);

                //Add filter conditions
                FilterExpression filterForBase = new FilterExpression();
                filterForBase.Conditions.Add(validFromConditionForBase);
                filterForBase.Conditions.Add(isActiveConditionForBase);
                filterForBase.Conditions.Add(isCopyConditionForBase);
                filterForBase.Conditions.Add(isApprovedConditionForBase);
               
                
                //Fetch Records that satisfy condition
                Entity baseprofile = new Entity("met_pricingprofile");
                QueryExpression queryForBase = new QueryExpression("met_pricingprofile") { ColumnSet = new ColumnSet(true) ,Criteria = filterForBase};
                EntityCollection retrievedAccount = serviceClient.RetrieveMultiple(queryForBase);
                
                //Setup for Multiple request execution
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

                foreach (var baseProfile in retrievedAccount.Entities)
                {
                   // Entity pricingProfile = new Entity("met_pricingprofileline");
                    var profileid = ((EntityReference)baseProfile.Attributes["met_originalprofileid"]).Id.ToString();
                    var profileidTodelete = new Guid(baseProfile.Attributes["met_pricingprofileid"].ToString());
                    Entity pricingProfile = new Entity("met_pricingprofile", new Guid(profileid));
                    pricingProfile["met_validfrom"] = (DateTime.Now);


                    Console.WriteLine("id: " + baseProfile.Id.ToString());
                    
                    var updateRequestForCreatingBase = new UpdateRequest()
                            {
                                Target = pricingProfile
                            };
                            request.Requests.Add(updateRequestForCreatingBase);
                    
                    /************
                     * Prepare filter conditions
                     * ** profile id associated = fetched record 
                     * ** is copy = Yes
                     * ** is copy updated = Yes
                     */
                    ConditionExpression profileIdForLine = new ConditionExpression() { AttributeName = "met_pricingprofileid", Operator = ConditionOperator.Equal };
                    ConditionExpression isCopyCondForLine = new ConditionExpression() { AttributeName = "met_iscopy", Operator = ConditionOperator.Equal };
                    ConditionExpression isCopyCondUpdatedForLine = new ConditionExpression() { AttributeName = "met_iscopyupdated", Operator = ConditionOperator.Equal };
                    profileIdForLine.Values.Add(new Guid(baseProfile.Attributes["met_pricingprofileid"].ToString()));
                    isCopyCondForLine.Values.Add(true);
                    isCopyCondUpdatedForLine.Values.Add(true);
                       
                    //Add filter conditions
                    FilterExpression filterForline = new FilterExpression();
                    filterForline.Conditions.Add(profileIdForLine);
                    filterForline.Conditions.Add(isCopyCondForLine);
                    
                    QueryExpression queryForLines = new QueryExpression("met_pricingprofileline") { ColumnSet = new ColumnSet(true), Criteria = filterForline };
                    EntityCollection retrievedProfileLines = serviceClient.RetrieveMultiple(queryForLines);

                    /******************
                     * Update in lines
                     ******************/
                    foreach (var lines in retrievedProfileLines.Entities)
                    {
                        //Create new profile lines 
                        if (lines.Attributes["met_iscopyupdated"].ToString().ToLower() == "true")
                        {
                            Entity newprofilelineCreate = new Entity("met_pricingprofileline");
                            
                            newprofilelineCreate["met_pricingprofileid"] = new EntityReference("met_pricingprofile", new Guid(profileid.ToString()));
                            newprofilelineCreate["met_validfrom"] = baseProfile.Attributes["met_validfrom"];
                            newprofilelineCreate["met_baseprofileid"] = lines.Attributes.ContainsKey("met_baseprofileid") && lines.Attributes["met_baseprofileid"] != null ? new EntityReference("met_pricingprofile", ((EntityReference)lines.Attributes["met_baseprofileid"]).Id) : null;
                            newprofilelineCreate["met_productid"] = lines.Attributes.ContainsKey("met_productid") && lines.Attributes["met_productid"] != null ? new EntityReference("product", ((EntityReference)lines.Attributes["met_productid"]).Id) : null;
                            newprofilelineCreate["met_productcategoryid"] = lines.Attributes.ContainsKey("met_productcategoryid") && lines.Attributes["met_productcategoryid"] != null ? new EntityReference("msdyn_productcategory", ((EntityReference)lines.Attributes["met_productcategoryid"]).Id) : null;
                            newprofilelineCreate["met_attachmenttypeid"] = lines.Attributes.ContainsKey("met_attachmenttypeid") && lines.Attributes["met_attachmenttypeid"] != null ? new EntityReference("met_attachmenttype", ((EntityReference)lines.Attributes["met_attachmenttypeid"]).Id) : null;
                            newprofilelineCreate["met_iscopy"] = true;
                            newprofilelineCreate["met_profilelinename"] = lines.Attributes["met_profilelinename"];
                            newprofilelineCreate["met_markuptype"] = lines.Attributes.ContainsKey("met_markuptype") ? lines.Attributes["met_markuptype"] : null;
                            newprofilelineCreate["met_earlierpercentage"] = lines.Attributes.ContainsKey("met_earlierpercentage") ? lines.Attributes["met_earlierpercentage"] : null; ;
                            newprofilelineCreate["exchangerate"] = lines.Attributes.ContainsKey("exchangerate") ? lines.Attributes["exchangerate"] : null;
                            newprofilelineCreate["met_percentage"] = lines.Attributes.ContainsKey("met_percentage") ? lines.Attributes["met_percentage"] : null;
                            newprofilelineCreate["met_wholesaleprice"] = lines.Attributes.ContainsKey("met_wholesaleprice") ? lines.Attributes["met_wholesaleprice"] : null;
                            newprofilelineCreate["met_srpprice"] = lines.Attributes.ContainsKey("met_srpprice") ? lines.Attributes["met_srpprice"] : null;
                            newprofilelineCreate["met_earliersrpprice"] = lines.Attributes.ContainsKey("met_earliersrpprice") ? lines.Attributes["met_earliersrpprice"] : null;
                            newprofilelineCreate["met_originalprofilelineid"] = new EntityReference("met_pricingprofile", new Guid(lines.Attributes["met_pricingprofilelineid"].ToString()));


                            var createRequest = new CreateRequest()
                            {
                                Target = newprofilelineCreate
                            };
                            request.Requests.Add(createRequest);


                            //update original profile line  valid to and move to inactive
                            Entity baseprofilelineCreate = new Entity("met_pricingprofileline", new Guid(((EntityReference)lines.Attributes["met_originalprofilelineid"]).Id.ToString()));
                            baseprofilelineCreate["met_validto"] = (DateTime.Now.Date.AddDays(-1));
                            baseprofilelineCreate.Attributes.Add("statecode", new OptionSetValue(1));
                            
                            var updateRequest = new UpdateRequest()
                            {
                                Target = baseprofilelineCreate
                            };
                            request.Requests.Add(updateRequest);

                        }
                        Console.WriteLine(lines.Attributes["met_pricingprofilelineid"].ToString());

                        //delete copy updated profile line
                        var deletepricingRequest = new DeleteRequest()
                        {
                            Target = new EntityReference("met_pricingprofileline", new Guid(lines.Attributes["met_pricingprofilelineid"].ToString()))
                        };
                        request.Requests.Add(deletepricingRequest);

                        if (request.Requests.Count >= 997)
                        {
                            BulkExecute(serviceClient, request);
                            request = BulkExecuteRequest();
                        }
                       

                    }
                    
                    var deleteRequest = new DeleteRequest()
                    {
                        Target = new EntityReference("met_pricingprofile", profileidTodelete)
                    };

                    request.Requests.Add(deleteRequest);
                    
                    
                }
                
                Console.WriteLine("Before execute------"+ request.Requests.Count);
                var response = new ExecuteMultipleResponse();
                if (request.Requests.Count <= 997)
                {
                    //BulkDeleteResponse(service, multipleRequest);
                     response = (ExecuteMultipleResponse)serviceClient.Execute(request);
                }
                Console.WriteLine("After execute");

                foreach (var r in response.Responses)
                {
                    if (r.Response != null)
                        Console.WriteLine("Success" + r.Response);
                    else if (r.Fault != null)
                        Console.WriteLine(r.Fault);
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("exception : "+ex.Message);
            }
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

        public static ExecuteMultipleResponse BulkExecute(ServiceClient service, ExecuteMultipleRequest multipleRequest)
        {
            ExecuteMultipleResponse multipleResponse = (ExecuteMultipleResponse)service.Execute(multipleRequest);

            return multipleResponse;
        }
    }
}
