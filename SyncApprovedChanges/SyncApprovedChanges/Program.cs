using System;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;
namespace SyncApprovedChanges
{
    internal class Program
    {
        static void Main()
        {
            
            const string clientId = "a9545b29-056e-4eab-8f31-911e9f24c52a";
            const string clientSecret = "HhXKcdJuY6Kid4EFdLhFBXmeIFHw2SuC";
            const string environment = "https://operations-metcash-pp-price.crm6.dynamics.com";
            var connectionString = $@"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";
            var serviceClient = new ServiceClient(connectionString);
            try
            {
                /********************************
                 * Prepare filter conditions
                 * ** validtime<= currenttime
                 * ** active record
                 * ** is copy = Yes
                 ********************************/
                ConditionExpression validFromConditionForBase = new ConditionExpression() { AttributeName = "met_validfrom", Operator = ConditionOperator.LessEqual };
                validFromConditionForBase.Values.Add(DateTime.Now);
                ConditionExpression isActiveConditionForBase = new ConditionExpression() { AttributeName = "met_iscopy", Operator= ConditionOperator.Equal };
                isActiveConditionForBase.Values.Add(true);
                ConditionExpression isCopyConditionForBase = new ConditionExpression() { AttributeName = "statecode", Operator = ConditionOperator.Equal }; ;
                isCopyConditionForBase.Values.Add(0);

                //Add filter conditions
                FilterExpression filterForBase = new FilterExpression();
                filterForBase.Conditions.Add(validFromConditionForBase);
                filterForBase.Conditions.Add(isActiveConditionForBase);
                filterForBase.Conditions.Add(isCopyConditionForBase);
               
                
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
                    Entity profileLine = new Entity("met_pricingprofileline");
                    var profileid = (new Guid(baseProfile.Attributes["met_pricingprofileid"].ToString()));
                    Console.WriteLine("id: " + profileid.ToString());

                    /************
                     * Prepare filter conditions
                     * ** profile id associated = fetched record 
                     * ** is copy = Yes
                     * ** is copy updated = Yes
                     */
                    ConditionExpression profileIdForLine = new ConditionExpression() { AttributeName = "met_pricingprofileid", Operator = ConditionOperator.Equal };
                    ConditionExpression isCopyCondForLine = new ConditionExpression() { AttributeName = "met_iscopyupdated", Operator = ConditionOperator.Equal };
                    ConditionExpression isCopyCondUpdatedForLine = new ConditionExpression() { AttributeName = "met_iscopy", Operator = ConditionOperator.Equal };
                    profileIdForLine.Values.Add(new Guid(baseProfile.Attributes["met_pricingprofileid"].ToString()));
                    isCopyCondForLine.Values.Add(true);
                    isCopyCondUpdatedForLine.Values.Add(true);
                    
                    //Add filter conditions
                    FilterExpression filterForline = new FilterExpression();
                    filterForline.Conditions.Add(profileIdForLine);
                    filterForline.Conditions.Add(isCopyCondForLine);
                    filterForline.Conditions.Add(isCopyCondUpdatedForLine);
                    QueryExpression queryForLines = new QueryExpression("met_pricingprofileline") { ColumnSet = new ColumnSet(true), Criteria = filterForline };

                    EntityCollection retrievedProfileLines = serviceClient.RetrieveMultiple(queryForLines);
                    
                    /******************
                     * Update in lines
                     ******************/
                    foreach (var lines in retrievedProfileLines.Entities)
                    {
                        Console.WriteLine("met_pricingprofilelineid: " + lines.Attributes["met_originalprofilelineid"]);
                        Entity baseprofilelineCreate = new Entity("met_pricingprofileline", new Guid(((EntityReference)lines.Attributes["met_originalprofilelineid"]).Id.ToString()));
                        //baseprofilelineCreate["met_pricingprofilelineid"] = new EntityReference("met_pricingprofileline", new Guid(((EntityReference)lines.Attributes["met_originalprofilelineid"]).Id.ToString()));
                        //baseprofilelineCreate["met_pricingprofilelineid"] =  new Guid(((EntityReference)lines.Attributes["met_originalprofilelineid"]).Id.ToString());

                        baseprofilelineCreate["met_profilelinename"] = lines.Attributes["met_profilelinename"];
                        //baseprofilelineCreate["met_pricingprofileid"] = new EntityReference("met_pricingprofile", new Guid(baseprofileGuid));
                        Console.WriteLine("met_pricingprofilelineid: " + lines.Attributes["met_pricingprofilelineid"]);
                        
                        var updateRequest = new UpdateRequest()
                        {
                            Target = baseprofilelineCreate
                        };
                        request.Requests.Add(updateRequest);
                        
                    }

                    //Make profile as inactive
                    Entity profileToInactive = new Entity("met_pricingprofile", profileid);
                    profileToInactive.Attributes.Add("statecode", new OptionSetValue(1));
                    profileToInactive.Attributes.Add("statuscode", new OptionSetValue(2));

                    var updateRequestToSetInactive = new UpdateRequest()
                    {
                        Target = profileToInactive
                    };
                    request.Requests.Add(updateRequestToSetInactive);
                }
                
                Console.WriteLine("Before execute");
                var response = (ExecuteMultipleResponse)serviceClient.Execute(request);
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
                Console.WriteLine("excepyion : "+ex.Message);
            }
        }
    }
}
