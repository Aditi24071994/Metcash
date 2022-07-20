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
using System.Linq;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;


namespace FunctionApp11
{
    public class Sample
    {
        public string DataFromDataverse { get; set; }
       

    }
    public static class Function1
    {
        [FunctionName("Function1")]

        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.System, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP tr3igger function processed a request.");

            /*-----------------------------------------
             * The next code will get the client id , client secret and environment url and connect to table testFY
             * testFy - crb3f_testfy , column considered - crb3f_datecustom
             * taking the top 10 count 
             * convert to csv 
             * save in local for reference
             */
            const string clientId = "a9545b29-056e-4eab-8f31-911e9f24c52a";
            const string clientSecret = "HhXKcdJuY6Kid4EFdLhFBXmeIFHw2SuC";
            const string environment = "https://operations-metcash-pp-price.crm6.dynamics.com";
            var connectionString = @$"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";
            using var serviceClient = new ServiceClient(connectionString);
            var accountsCollection = await serviceClient.RetrieveMultipleAsync(new QueryExpression("account")
            {
                ColumnSet = new ColumnSet("createdon"),
                TopCount = 10
            });
           
            Console.WriteLine(string.Join("\n",
                accountsCollection.Entities
                    .Select(x => $"{x.GetAttributeValue<DateTime>("createdon")}, {x.Id}")));

           
            var DataToUpload = new List<Sample> {
               new Sample{ DataFromDataverse =string.Join("\n",
                accountsCollection.Entities
                    .Select(x => $"{x.GetAttributeValue<DateTime>("createdon")}, {x.Id}")) }
            };

            string properties = string.Join(",", DataToUpload[0].GetType().GetProperties().Select(i => i.Name));

            var info = from member in DataToUpload
                       let record = string.Join(",", member.GetType().GetProperties().Select(j => j.GetValue(member)))
                       select record;

            var csv = new List<string>
            {
                properties
            };

            csv.AddRange(info);

            string newPath = Directory.GetCurrentDirectory() + "/Data";

            //Check if the directory exists
            //
            if (Directory.Exists(newPath))
                Directory.Delete(newPath, true);

            Directory.CreateDirectory(newPath);

            File.WriteAllLines(newPath + "/FileToBeUploaded.csv", csv);

            /*-----------------------------------------
             * The next code will get the storage account details and directly create a file there
             * storage account - dataverseconnect0713
            */
            try
            {

                CloudStorageAccount account = CloudStorageAccount.Parse("DefaultEndpointsProtocol=https;AccountName=test07132022;AccountKey=uocw/VVOwhS5A65+o8Y59/u6NElxSzc1aE/iu/JBSdNU/Bg529D2efJRPvWIyt/eGtexxCQQ3DDH+ASt7oDtWQ==;EndpointSuffix=core.windows.net");

                CloudBlobClient client = account.CreateCloudBlobClient();

                CloudBlobContainer container = client.GetContainerReference("test");

                await container.CreateIfNotExistsAsync();

                CloudBlockBlob blob = container.GetBlockBlobReference("data.csv");
               
                using (CloudBlobStream x = blob.OpenWriteAsync().Result)
                {
                    foreach (var rec in csv)
                    {
                        x.Write(System.Text.Encoding.Default.GetBytes(rec.ToString() + "\n"));
                    }
                    x.Flush();
                    x.Close();
                }
                

                Console.WriteLine("File Uploaded!");

            }
            catch (Exception ex)
            {
                Console.WriteLine("File could not be uploaded due to: " + ex.Message);
            }
            return new OkObjectResult(string.Join("\n",
                accountsCollection.Entities
                    .Select(x => $"{x.GetAttributeValue<DateTime>("createdon")}, {x.Id}")));


            /*return new OkObjectResult(string.Join("\n",
                accountsCollection.Entities
                    .Select(x => $"{x.GetAttributeValue<string>("name")}, {x.Id}")));*/
        }
    }
}
