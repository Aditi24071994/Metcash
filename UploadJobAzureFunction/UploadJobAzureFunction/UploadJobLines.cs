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
    public static class UploadJobLines
    {
        [FunctionName("UploadJobLines")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest headerRequest,
            ILogger log)
        {
            log.LogInformation("UploadJobLines Function triggered");

            string id = headerRequest.Query["id"];
            string type = headerRequest.Query["type"];
            int countOfCreated = 0, countOfUpdated = 0;

            var clientId = "a9545b29-056e-4eab-8f31-911e9f24c52a";
            const string clientSecret = "HhXKcdJuY6Kid4EFdLhFBXmeIFHw2SuC";
            const string environment = "https://operations-metcash-pp-price.crm6.dynamics.com";
            var connectionString = @$"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";

            var serviceClient = new ServiceClient(connectionString);
            var fetchxmlWow = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>  <entity name='met_pricingprofile'>'<attribute name='met_pricingprofileid' />'<order attribute='met_profilenumber' descending='false' /><link-entity name='account' from='accountid' to='met_customerid' link-type='inner' alias='ap'>' <filter type='and'>'   <condition attribute='name' operator='eq' value='Woolworths Group Ltd' />' </filter>'</link-entity>'<link-entity name='met_pricingprofiletype' from='met_pricingprofiletypeid' to='met_profiletypeid' link-type='inner' alias='aq'>' <filter type='and'>'   <condition attribute='met_profiletypename' operator='eq' value='Raw' />' </filter>'</link-entity>  </entity></fetch>";
            var fetchxmlCol = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>  <entity name='met_pricingprofile'>'<attribute name='met_pricingprofileid' />'<order attribute='met_profilenumber' descending='false' /><link-entity name='account' from='accountid' to='met_customerid' link-type='inner' alias='ap'>' <filter type='and'>'   <condition attribute='name' operator='eq' value='Coles Group Ltd' />' </filter>'</link-entity>'<link-entity name='met_pricingprofiletype' from='met_pricingprofiletypeid' to='met_profiletypeid' link-type='inner' alias='aq'>' <filter type='and'>'   <condition attribute='met_profiletypename' operator='eq' value='Raw' />' </filter>'</link-entity>  </entity></fetch>";
            var fetchxmlForProfileLine = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>  <entity name='met_pricingprofileline'><attribute name='met_pricingprofilelineid' /><attribute name='met_pricingprofileid' /><attribute name='met_productid' /><order attribute='met_profilelinename' descending='false' /><link-entity name='product' from='productid' to='met_productid' link-type='inner' alias='ac' />  </entity></fetch>";

            EntityCollection colMasterProfileLines = PopulateMasters("met_pricingprofileline", serviceClient, fetchxmlForProfileLine);
            EntityCollection colMasterProduct = PopulateMasters("product", serviceClient,"");
            EntityCollection colsMasterPricingProfile = PopulateMasters("met_pricingprofile", serviceClient, fetchxmlCol);
            EntityCollection wowMasterPricingProfile = PopulateMasters("met_pricingprofile", serviceClient, fetchxmlWow);

            Console.WriteLine("in above " + colMasterProfileLines.Entities[0].Attributes["met_productid"]);

            /* ********************************
             * Code to download the file fromjob
             * *********************************/
            InitializeFileBlocksDownloadRequest initializeFile = new()
            {
                FileAttributeName = "met_uploadfile", // attribute name
                Target = new EntityReference("met_uploadjob", Guid.Parse(id))
            };
            InitializeFileBlocksDownloadResponse initializeFileResponse = (InitializeFileBlocksDownloadResponse)serviceClient.Execute(initializeFile);

            var fileContinuationToken = initializeFileResponse.FileContinuationToken;
            // code to downlod the file.
            DownloadBlockRequest downloadRequest = new()
            {
                //Offset = 0,
                //BlockLength = (long)4 * 1024 * 1024,
                FileContinuationToken = fileContinuationToken
            };
            DownloadBlockResponse downloadBlockResponse = (DownloadBlockResponse)serviceClient.Execute(downloadRequest);
            byte[] byteArray = downloadBlockResponse.Data;
            MemoryStream stream = new(byteArray);
            //tracing.Trace("stream successfully");

            SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = document.WorkbookPart;
            WorksheetPart worksheetPart = GetWorksheetPart(workbookPart, "BASE CHANGES");
            System.Globalization.CultureInfo provider = System.Globalization.CultureInfo.InvariantCulture;
            int rowcount = worksheetPart.Worksheet.Descendants<Row>().Count();
            Worksheet workSheet = worksheetPart.Worksheet;
            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();
            Console.WriteLine(rows.Count().ToString());
            DataTable dt = new();
            DataTable datatableWithRecords = new();
            datatableWithRecords.Columns.Add("Code");
            datatableWithRecords.Columns.Add("Competitor");
            datatableWithRecords.Columns.Add("Subrange");
            datatableWithRecords.Columns.Add("SRP Old");
            datatableWithRecords.Columns.Add("SRP New");
            datatableWithRecords.Columns.Add("Description");
            datatableWithRecords.Columns.Add("Error Message");
            List<string> rowsFromJobFile = new();
            foreach (Cell cell in rows.ElementAt(0))
            {
                _ = dt.Columns.Add(GetCellValue(document, cell));
            }

            //fetch the header row index values

            foreach (Row row in rows) //this will also include your header row...
            {
                DataRow tempRow = dt.NewRow();

                for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                {
                    //Console.WriteLine("inside");
                    try
                    {

                        tempRow[i] = GetCellValue(document, row.Descendants<Cell>().ElementAt(i));

                        rowsFromJobFile.Add((string)tempRow[i]);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                break;

                //dt.Rows.Add(tempRow);


            }

            //Initialize the columns
            int colPosWOWPBSrp = rowsFromJobFile.IndexOf("SRP WOW PB");
            int colPosWOWFRUSrp = rowsFromJobFile.IndexOf("SRP WOW FRU");
            int colPosWOWLPBLPED = rowsFromJobFile.IndexOf("LPED WOW PB");
            int colPosWOWLFRULPED = rowsFromJobFile.IndexOf("LPED WOW FRU");
            int colPosCOLPBSrp = rowsFromJobFile.IndexOf("SRP COL PB");
            int colPosCOLFRUSrp = rowsFromJobFile.IndexOf("SRP COL FRU");
            int colPosCode = rowsFromJobFile.IndexOf("CODE");
            int colPosDesc = rowsFromJobFile.IndexOf("DESCRIPTION");
            int colPosSubrange = rowsFromJobFile.IndexOf("SUBRANGE");
            int counter = 0;
            int counterForErrors = 0;
            Console.WriteLine(colPosCOLFRUSrp);

            foreach (Row row in rows) //this will also include your header row...
            {
                if (counter == 0)
                {

                }
                else
                {
                    DataRow tempRow = datatableWithRecords.NewRow();
                    
                  tempRow["SRP Old"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosWOWPBSrp));
                  tempRow["SRP New"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosWOWFRUSrp));
                  tempRow["Description"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosDesc));
                  tempRow["Code"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosCode));
                  tempRow["Subrange"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosSubrange));
                  tempRow["Competitor"] = "WOW";
                  datatableWithRecords.Rows.Add(tempRow);
                    /*
                  tempRow = datatableWithRecords.NewRow();
                  tempRow["SRP Old"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosWOWLPBLPED));
                  tempRow["SRP New"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosWOWLFRULPED));
                  tempRow["Description"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosDesc));
                  tempRow["Code"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosCode));
                  tempRow["Subrange"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosSubrange));
                  tempRow["Competitor"] = "WOW LPED";
                  datatableWithRecords.Rows.Add(tempRow);
                  */

                    tempRow = datatableWithRecords.NewRow();
                    tempRow["SRP Old"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosCOLPBSrp));
                    tempRow["SRP New"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosCOLFRUSrp));
                    tempRow["Description"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosDesc));
                    tempRow["Code"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosCode));
                    tempRow["Subrange"] = GetCellValue(document, row.Descendants<Cell>().ElementAt(colPosSubrange));
                    tempRow["Competitor"] = "COLS";
                    datatableWithRecords.Rows.Add(tempRow);

                }

                counter++;
            }


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

            var itemIDForProduct = new Guid();
            var itemIDForPricingProfile = new Guid();
            Entity joblineAction = new();
            Console.WriteLine(datatableWithRecords.Rows.Count.ToString());
            
            var filter = datatableWithRecords.AsEnumerable().
                       Where(x => x.Field<string>("Subrange") != "");
            datatableWithRecords = datatableWithRecords.AsEnumerable().
                       Where(x => x.Field<string>("SRP New") != "").CopyToDataTable();
            var query = (from row in filter.AsEnumerable()
                         
                         group row by new
                         {
                             Subrange = row.Field<string>("Subrange"),
                             Competitor = row.Field<string>("Competitor"),
                             Price = row.Field<string>("SRP New")
                         }

                         into grp
                         select new
                         {
                             Subrange = grp.Key.Subrange,
                             Competitor = grp.Key.Competitor,
                             Price = grp.Key.Price
                         }).ToList();

            Console.WriteLine(query.Count);




            foreach (DataRow rowsForDt in datatableWithRecords.Rows)
            {
                var errorMessage = "";
                var jobLine = new Entity("met_uploadjobprofileline");
                var item = colMasterProduct.Entities.Where(x => x.Attributes["msdyn_productnumber"].ToString().ToLower() == rowsForDt["Code"].ToString().ToLower()).FirstOrDefault();
                var rowCountForFetch = query.Where(x => x.Competitor.ToString().ToLower() == rowsForDt["Competitor"].ToString().ToLower()
                && x.Subrange.ToString().ToLower() == rowsForDt["Subrange"].ToString().ToLower()).Count();
                var itemFromPRofileLines = colMasterProfileLines.Entities.Where(x => x.Attributes["met_productid"].ToString().ToLower() == rowsForDt["Code"].ToString().ToLower()).FirstOrDefault();

                if (rowsForDt["Competitor"].ToString().Contains("WOW")){

                    jobLine["met_pricingprofileid"] = new EntityReference("met_pricingprofile", wowMasterPricingProfile.Entities.FirstOrDefault().Id);
                    itemIDForPricingProfile = wowMasterPricingProfile.Entities.FirstOrDefault().Id;
                }
                else
                {
                    jobLine["met_pricingprofileid"] = new EntityReference("met_pricingprofile", colsMasterPricingProfile.Entities.FirstOrDefault().Id);

                    itemIDForPricingProfile = colsMasterPricingProfile.Entities.FirstOrDefault().Id;

                }



                if (item == null)
                {
                    errorMessage = errorMessage + " | " + "Item "+ rowsForDt["Code"].ToString() + " does not exist";
                    jobLine["statuscode"] = new OptionSetValue(862200002);
                    jobLine["met_newsrp"] = new Money(Convert.ToDecimal(rowsForDt["SRP New"].ToString()));
                    counterForErrors++;
                }
                else if (rowCountForFetch > 1)
                {
                    errorMessage = errorMessage + " | " + "Subrange issue";
                    itemIDForProduct = item.Id;
                    jobLine["met_itemid"] = new EntityReference("product", itemIDForProduct);
                    jobLine["statuscode"] = new OptionSetValue(862200001);
                    jobLine["met_newsrp"] = new Money(Convert.ToDecimal(rowsForDt["SRP New"].ToString()));
                    counterForErrors++;

                }
                else if(((Convert.ToDouble(rowsForDt["SRP New"])- (Convert.ToDouble(rowsForDt["SRP Old"]))) / Convert.ToDouble(rowsForDt["SRP Old"])) >0.4)
                {
                    errorMessage = errorMessage + " | " + "Variance > 40%";
                    itemIDForProduct = item.Id;
                    jobLine["met_itemid"] = new EntityReference("product", itemIDForProduct);
                    jobLine["statuscode"] = new OptionSetValue(862200001);
                    jobLine["met_newsrp"] = new Money(Convert.ToDecimal(rowsForDt["SRP New"].ToString()));
                    counterForErrors++;

                }
                else
                {
                    itemIDForProduct = item.Id;
                    jobLine["met_itemid"] = new EntityReference("product", itemIDForProduct);
                    jobLine["statuscode"] = new OptionSetValue(862200000);
                    jobLine["met_newsrp"] = new Money(Convert.ToDecimal(rowsForDt["SRP New"].ToString()));

                }
                Console.WriteLine(errorMessage);
                var intEntity = colMasterProfileLines.Entities.Where(x => x.Attributes["met_productid"] != null);
                var itemForLine = intEntity.Where(x => x.Attributes["met_productid"] != null && x.Attributes["met_productid"].ToString().ToLower() == itemIDForProduct.ToString() &&
                x.Attributes["met_pricingprofileid"].ToString() == itemIDForPricingProfile.ToString()).FirstOrDefault();
                
                if(itemForLine != null)
                {
                    jobLine["met_existingprofilelineid"] = new EntityReference("product", itemForLine.Id); ;
                }

                jobLine["met_errormessage"] = errorMessage;
                jobLine["met_uploadjobid"] = new EntityReference("met_uploadjob", new Guid(id));
                var jobLineCreate = new CreateRequest()
                {
                    Target = jobLine
                };
                requestForMultipleFetch.Requests.Add(jobLineCreate);


            }

            var jobUpload = new Entity("met_uploadjob",new Guid(id));
            if (counterForErrors > 0)
            {
                jobUpload["statuscode"] = new OptionSetValue(862200001);
            }
            else
            {
                jobUpload["statuscode"] = new OptionSetValue(862200000);
            }
            var jobUploadCreate = new UpdateRequest()
            {
                Target = jobUpload
            };
            requestForMultipleFetch.Requests.Add(jobUploadCreate);

            Console.WriteLine("Before execute");
             var response = (ExecuteMultipleResponse)serviceClient.Execute(requestForMultipleFetch);
             Console.WriteLine("After execute");
            /*
             foreach (var r in response.Responses)
             {
                 if (r.Response != null)


                     Console.WriteLine("Success" + r.Response+"----"+r.GetType().Name);
                 else if (r.Fault != null)
                     Console.WriteLine(r.Fault);
             }

            */
            return null;
        }

        private static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {

            Console.WriteLine("The File is in expected format " + sheetName);
            string relId = workbookPart.Workbook.Descendants<Sheet>().First(s => sheetName.Equals(s.Name)).Id;
            Console.WriteLine("The File is in expected format " + relId);

            return (WorksheetPart)workbookPart.GetPartById(relId);
        }


        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = "";
            try
            {
                if (cell.CellValue != null)
                {
                    value = cell.CellValue.InnerXml;
                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                    }
                    else
                    {
                        return value;
                    }
                }
                else
                {
                    return value;
                }
            }
            catch (Exception ex)
            {
                return value;
            }
        }
        public static EntityCollection PopulateMasters(string name, ServiceClient service,String fetchString)
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