//----------------------------------------------------------------------------------
// Microsoft Developer & Platform Evangelism
//
// Copyright (c) Microsoft Corporation. All rights reserved.
//
// THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
// EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES 
// OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
//----------------------------------------------------------------------------------
// The example companies, organizations, products, domain names,
// e-mail addresses, logos, people, places, and events depicted
// herein are fictitious.  No association with any real company,
// organization, product, domain name, email address, logo, person,
// places, or events is intended or should be inferred.
//----------------------------------------------------------------------------------

namespace ExcelAutomationAzureAspose
{
    using Microsoft.WindowsAzure;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.Blob;
    using System;

    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    // P.R Add reference to Aspose trial version.
    using Aspose.Cells;

    /// <summary>


    /// </summary>
    public class Program
    {
         static void Main(string[] args)
        {
            Console.WriteLine("Demonstrating 3rd party excel automation using Azure Blob Storage ");
            // P.R Initialize connection objects in Azure
            string sInputContainer = "exceltemplates";// Replace these with container names
            string sOutputContainer = "excelresults";
               
            string a = CloudConfigurationManager.GetSetting("StorageConnectionString");
            CloudStorageAccount storageAccount = CreateStorageAccountFromConnectionString(CloudConfigurationManager.GetSetting("StorageConnectionString"));
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
            CloudBlobContainer container = blobClient.GetContainerReference(sInputContainer);
            CloudBlobContainer containerOut = blobClient.GetContainerReference(sOutputContainer);
            try
            {
                container.CreateIfNotExistsAsync();
            }
            catch (StorageException e)
            {
                throw e;
            }      
            String fstreamPath = DownloadFileStreamFromAzure(container, "Sample.xlsm");
           
            Console.WriteLine("File Stream downloaded From Azure BLOB storage");
            FileStream fstream = new FileStream(fstreamPath, FileMode.Open);

            Workbook wbAspose = LoadAsposeWorkbookFromStream(fstream);

            Console.WriteLine("ASPose workbook loaded from File Stream");

            String fsStreamOutput= ManipulateAsposeWorksheet(wbAspose);

            bool bUploaded = UploadFileStreamToAzure(containerOut, "Sample.xlsm", fsStreamOutput);

            Console.WriteLine("ASPose workbook uploaded to Azure BLOB storage");

            fstream.Close();
            Console.WriteLine("Press any key to exit");
            Console.ReadLine();
        }

        private static string DownloadFileStreamFromAzure(CloudBlobContainer container, string blobName)
        {

            CloudBlockBlob blockBlob = container.GetBlockBlobReference(blobName);
            
            using (var fileStream = File.OpenWrite(Path.Combine(Path.GetTempPath(), "Temp.xlsm")))
            {
                blockBlob.DownloadToStream(fileStream);
                return Path.Combine(Path.GetTempPath()+ "Temp.xlsm");
            }

        }

        private static bool UploadFileStreamToAzure(CloudBlobContainer container, string blobName, string streamPath)
        {

            CloudBlockBlob blockBlob = container.GetBlockBlobReference(blobName);            
            using (var fileStream = File.OpenRead(streamPath))
            {
                blockBlob.UploadFromStream(fileStream);
                return true;
            }
        }

        private static Workbook LoadAsposeWorkbookFromStream(FileStream fstream)
        {
           Workbook workbook2 = new Workbook(fstream);

            return workbook2;
        }

        private static string ManipulateAsposeWorksheet(Workbook workbookStream)
        {
            Worksheet worksheet = workbookStream.Worksheets[0];
            worksheet.Cells["R1"].PutValue(1);
            worksheet.Cells["R2"].PutValue(2);
            worksheet.Cells["R3"].PutValue(3);
            worksheet.Cells["R4"].Formula = "=SUM(R1:R3)";
            workbookStream.CalculateFormula();
            string value = worksheet.Cells["R4"].Value.ToString();
            Console.WriteLine("Sum calculated by ASPose is " + value);
            

            using (var fileStream = File.OpenWrite(Path.Combine(Path.GetTempPath(), "TempOut.xlsm")))
            {
                workbookStream.Save(fileStream, SaveFormat.Xlsm);
                Console.WriteLine("Output stream saved as XLSM ");
                return Path.Combine(Path.GetTempPath() + "TempOut.xlsm");
            }
            //return Path.Combine(Path.GetTempPath() + "TempOut.xlsm");
        }
               
        private static CloudStorageAccount CreateStorageAccountFromConnectionString(string storageConnectionString)
        {
            CloudStorageAccount storageAccount;
            try
            {
                storageAccount = CloudStorageAccount.Parse(storageConnectionString);
            }
            catch (FormatException)
            {
                Console.WriteLine("Invalid storage account information provided. Please confirm the AccountName and AccountKey are valid in the app.config file - then restart the sample.");
                Console.ReadLine();
                throw;
            }
            catch (ArgumentException)
            {
                Console.WriteLine("Invalid storage account information provided. Please confirm the AccountName and AccountKey are valid in the app.config file - then restart the sample.");
                Console.ReadLine();
                throw;
            }

            return storageAccount;
        }

    }
}
