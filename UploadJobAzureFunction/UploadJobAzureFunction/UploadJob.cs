using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.Runtime.Serialization;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Collections.Generic;
using System.Data;


namespace UploadJobAzureFunction
{
    public static class UploadJob
    {
        [FunctionName("UploadJob")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest headerRequest,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            int countOfErrors = 0;
            string id = headerRequest.Query["id"];

            var clientId = "a9545b29-056e-4eab-8f31-911e9f24c52a";
            const string clientSecret = "HhXKcdJuY6Kid4EFdLhFBXmeIFHw2SuC";
            const string environment = "https://operations-metcash-pp-price.crm6.dynamics.com";
            var connectionString = @$"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";
            
            var serviceClient = new ServiceClient(connectionString);

            ConditionExpression conditionForTemplate = new() { AttributeName = "met_name", Operator = ConditionOperator.Equal};
            conditionForTemplate.Values.Add("Retail Pricing");
            FilterExpression filterForTemplate = new();
            filterForTemplate.Conditions.Add(conditionForTemplate);


            QueryExpression query = new("met_uploadtemplate")
            {
                ColumnSet = new ColumnSet(true)
            };
            query.Criteria.AddFilter(filterForTemplate);
            EntityCollection resultFromTemp = serviceClient.RetrieveMultiple(query);


            log.LogInformation("met_uploadjob : " + resultFromTemp.Entities[0].Attributes["met_file"]);

            /* ********************************
             * Code to download the file fromjob
             * *********************************/
            InitializeFileBlocksDownloadRequest initializeFile = new()
            {
                FileAttributeName = "met_uploadfile", // attribute name
                Target = new EntityReference("met_uploadjob", Guid.Parse(id))
            };
            InitializeFileBlocksDownloadResponse initializeFileResponse = (InitializeFileBlocksDownloadResponse)serviceClient.Execute(initializeFile);
            log.LogInformation($"File Name: {initializeFileResponse.FileName}");
            log.LogInformation($"File size (bytes): {initializeFileResponse.FileSizeInBytes}");
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
            CultureInfo provider = CultureInfo.InvariantCulture;
            int rowcount = worksheetPart.Worksheet.Descendants<Row>().Count();
            Worksheet workSheet = worksheetPart.Worksheet;
            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();
            Console.WriteLine(rows.Count().ToString());
            DataTable dt = new();
            List<string> arrCell = new();
            foreach (Cell cell in rows.ElementAt(0))
            {
                _ = dt.Columns.Add(GetCellValue(document, cell));
            }

            foreach (Row row in rows) //this will also include your header row...
            {
                DataRow tempRow = dt.NewRow();

                for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                {
                    //Console.WriteLine("inside");
                    try
                    {
                        
                        tempRow[i] = GetCellValue(document, row.Descendants<Cell>().ElementAt(i));
                        Console.WriteLine("---- actual file------"+tempRow[i]);
                        arrCell.Add((string)tempRow[i]);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                dt.Rows.Add(tempRow);

                break;
            }
            /* ********************************
             * Code to download the file from Template
             * *********************************/
            InitializeFileBlocksDownloadRequest initializeTempFile = new()
            {
                FileAttributeName = "met_file", // attribute name
                Target = new EntityReference("met_uploadtemplate", new Guid((resultFromTemp.Entities[0].Attributes["met_uploadtemplateid"]).ToString()))
            };
            InitializeFileBlocksDownloadResponse initializeTempFileResponse = (InitializeFileBlocksDownloadResponse)serviceClient.Execute(initializeTempFile);
            log.LogInformation($"File Name: {initializeTempFileResponse.FileName}");
            log.LogInformation($"File size (bytes): {initializeTempFileResponse.FileSizeInBytes}");

            if (!(new FileInfo(initializeTempFileResponse.FileName).Extension == new FileInfo(initializeFileResponse.FileName).Extension))
            {
                log.LogInformation("The File is not in expected format ");
            }
            else
            {
                log.LogInformation("The File is in expected format ");
                var fileContTempToken = initializeTempFileResponse.FileContinuationToken;
                // code to downlod the file.
                DownloadBlockRequest downloadTemplateRequest = new()
                {
                    //Offset = 0,
                    //BlockLength = (long)4 * 1024 * 1024,
                    FileContinuationToken = fileContTempToken
                };
                DownloadBlockResponse downloadTempBlockResponse = (DownloadBlockResponse)serviceClient.Execute(downloadTemplateRequest);
                byte[] byteArrayTemplate = downloadTempBlockResponse.Data;
                MemoryStream streamForTemplate = new(byteArrayTemplate);
                SpreadsheetDocument documentForTemplate = SpreadsheetDocument.Open(streamForTemplate, false);
                WorkbookPart workbookPartForTemplate = documentForTemplate.WorkbookPart;
                WorksheetPart worksheetPartForTemplate = GetWorksheetPart(workbookPartForTemplate, "BASE CHANGES");
                log.LogInformation("The Fisle is in expected format ");
                CultureInfo providerForTemplate = CultureInfo.InvariantCulture;
                int rowcountForTemplate = worksheetPartForTemplate.Worksheet.Descendants<Row>().Count();
                Worksheet workSheetForTemplate = worksheetPartForTemplate.Worksheet;
                SheetData sheetDataForTemplate = workSheetForTemplate.GetFirstChild<SheetData>();
                IEnumerable<Row> rowsForTemplate = sheetDataForTemplate.Descendants<Row>();
                Console.WriteLine(rowsForTemplate.Count().ToString());
                DataTable dtForTemplate = new DataTable();
                foreach (Cell cell in rowsForTemplate.ElementAt(0))
                {
                    _ = dtForTemplate.Columns.Add(GetCellValue(document, cell));
                }
                var arrCellForTemplate = Array.Empty<string>();
                foreach (Row row in rowsForTemplate) //this will also include your header row...
                {
                    DataRow tempRow1 = dtForTemplate.NewRow();

                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        //Console.WriteLine("inside");
                        try
                        {
                           
                            tempRow1[i] = GetCellValue(documentForTemplate, row.Descendants<Cell>().ElementAt(i));
                            Console.WriteLine("----template row----"+tempRow1[i]);

                            if (!arrCell.Contains((string)tempRow1[i]))
                            {
                                countOfErrors++;
                            }
                            //arrCellForTemplate.Append(tempRow[i]);

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }

                    dtForTemplate.Rows.Add(tempRow1);

                    break;
                }

               


                log.LogInformation("Count of errors : " + countOfErrors);
            }

            return null;
        }

        private static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {

            Console.WriteLine("The File is in expected format "+sheetName);
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
            catch(Exception ex)
            {
                return value;
            }
        }
    }
}