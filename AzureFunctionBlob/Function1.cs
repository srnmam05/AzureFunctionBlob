using System;
using System.IO;
using AzureFunctionBlob.Service;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;

namespace AzureFunctionBlob
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static void Run([BlobTrigger("samples-workitems/{name}", Connection = "")]Stream myBlob, string name, ILogger log)
        {
            var service = new SaveToDb();
            var excelNames = service.ExcelName(name);
            service.LicneseListPrice(myBlob);
            foreach (var excelName in excelNames)
            {
                log.LogInformation($" Excel:{excelName.ExcelName} 存到資料庫成功");
            }
        }
    }
}
