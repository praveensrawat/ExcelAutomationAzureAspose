
# ExcelAutomationAzureAspose
Excel Automation using ASPose 3rd party component and Azure Blob Storage.

1. Create a storage account in Azure. 

2. Create an input and output BLOB service container.

3.Replace account and keys with your credentials

4. Add a NuGet reference to a trial version of ASPose Cells. 

5. Upload a XSLM file to the input container.

The sample will download stream from the input container, append numbers to the first worksheet and perform a summation. 

The result will be stored into the output container.

