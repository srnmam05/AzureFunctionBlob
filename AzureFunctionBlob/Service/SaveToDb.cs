using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace AzureFunctionBlob.Service
{
    /// <summary>
    /// 資料匯入
    /// </summary>
    public class SaveToDb : BaseService
    {
        /// <summary>
        /// 取得 Excel 頁簽
        /// </summary>
        /// <param name="blobStream"></param>
        /// <param name="Sheet"></param>
        /// <returns></returns>
        private (ISheet TWD, ISheet USD) GetSheet(Stream blobStream)
        {
            var file = new XSSFWorkbook(blobStream);
            ISheet sheetTWD = file.GetSheet("TWD");
            ISheet sheetUSD = file.GetSheet("USD");
            return (sheetTWD, sheetUSD);
        }
        /// <summary>
        /// 儲存多筆上傳檔案時的名稱
        /// </summary>
        /// <param name="Name"></param>
        /// <returns></returns>
        public List<DbViewModel> ExcelName(string Name)
        {
            List<DbViewModel> ExcelName = new List<DbViewModel>();
            var name = new DbViewModel()
            {
                ExcelName = Name
            };
            ExcelName.Add(name);
            return ExcelName;
        }
        /// <summary>
        /// 取得相對應 USD
        /// </summary>
        /// <param name="data"></param>
        /// <param name="OfferID"></param>
        /// <returns></returns>
        private (decimal ListPrice, decimal ErpPrice) USDPrice(List<DbViewModel> data, string OfferID)
        {
            decimal listPrice = 0;
            decimal erpPrice = 0;
            foreach (var usd in data)
            {
                if (OfferID == usd.OfferID)
                {
                    listPrice = usd.ListPrice_TWD;
                    erpPrice = usd.ERPPrice_TWD;
                    break;
                }
            }
            return (listPrice, erpPrice);
        }
        /// <summary>
        /// 存進資料庫的方法
        /// </summary>
        /// <param name="blobStream"></param>
        public void LicneseListPrice(Stream blobStream)
        {
            var sheet = GetSheet(blobStream);
            List<DbViewModel> ListPriceTWD = GetData(sheet.TWD);
            List<DbPriceViewModel> ListPriceUSD = GetData(sheet.USD);
            DeleteData(ListPriceTWD);
            foreach (var data in ListPriceTWD)
            {
                try
                {
                    var checkStartDate = data.ValidFromDate.AddMonths(-1);
                    var checkEndDate = data.ValidFromDate.AddMonths(1);
                    var USD = USDPrice(ListPriceUSD, data.OfferID);
                    data.ListPrice_USD = USD.ListPrice;
                    data.ERPPrice_USD = USD.ErpPrice;
                    var result = new DbContext()
                    {
                        Id = Guid.NewGuid(),
                        LicenseType = data.LicenseType,
                        ValidFromDate = data.ValidFromDate,
                        ValidToDate = data.ValidToDate,
                        OfferDisplayName = data.OfferDisplayName,
                        OfferId = data.OfferID,
                        LicenseAgreement = data.LicenseAgreement,
                        PurchaseUnit = data.PurchaseUnit,
                        SecondaryLicenseType = data.SecondaryLicenseType,
                        EndCustomerType = data.EndCustomerType,
                        ListPriceTwd = data.ListPrice_TWD,
                        ErppriceTwd = data.ERPPrice_TWD,
                        Material = data.Material,
                        ListPriceUsd = data.ListPrice_USD,
                        ErppriceUsd = data.ERPPrice_USD,
                    };
                    db.DbContext.Add(result);
                    db.SaveChanges();
                    Console.WriteLine($"Add {result.OfferDisplayName}-{result.OfferId}-{result.ValidFromDate} Success");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{data.OfferDisplayName}-{data.OfferID}-{ex.Message}");
                }
            }
        }
        /// <summary>
        /// 取得 Excel 內資料
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private List<DbViewModel> GetData(ISheet sheet)
        {
            List<DbViewModel> ListPrice = new List<DbViewModel>();
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                DataFormatter formatter = new DataFormatter();
                String listPrice = formatter.FormatCellValue(sheet.GetRow(row).GetCell(9));
                String ERPPrice = formatter.FormatCellValue(sheet.GetRow(row).GetCell(10));
                var data = new DbViewModel()
                {
                    LicenseType = sheet.GetRow(row).GetCell(0).StringCellValue,
                    ValidFromDate = DateTime.ParseExact(sheet.GetRow(row).GetCell(1).StringCellValue, "yyyyMMdd", null, System.Globalization.DateTimeStyles.AllowWhiteSpaces),
                    ValidToDate = DateTime.ParseExact(sheet.GetRow(row).GetCell(2).StringCellValue, "yyyyMMdd", null, System.Globalization.DateTimeStyles.AllowWhiteSpaces),
                    OfferDisplayName = sheet.GetRow(row).GetCell(3).StringCellValue,
                    OfferID = sheet.GetRow(row).GetCell(4).StringCellValue,
                    LicenseAgreement = sheet.GetRow(row).GetCell(5).StringCellValue,
                    PurchaseUnit = sheet.GetRow(row).GetCell(6).StringCellValue,
                    SecondaryLicenseType = sheet.GetRow(row).GetCell(7).StringCellValue,
                    EndCustomerType = sheet.GetRow(row).GetCell(8).StringCellValue,
                    ListPrice_TWD = decimal.Parse(listPrice),
                    ERPPrice_TWD = decimal.Parse(ERPPrice),
                    Material = sheet.GetRow(row).GetCell(11).StringCellValue,
                };
                ListPrice.Add(data);
            }
            return ListPrice;
        }
        /// <summary>
        /// 如果上傳的 Excel 是資料庫最新日期的資料，刪除資料庫此日期資料
        /// </summary>
        /// <param name="ListPrice"></param>
        private void DeleteData(List<DbViewModel> ListPrice)
        {
            if (db.DbContext.FirstOrDefault() != null)
            {
                var date = db.DbContext.OrderByDescending(X => X.ValidFromDate).FirstOrDefault().ValidFromDate;
                if (date == ListPrice.FirstOrDefault().ValidFromDate)
                {
                    var data = db.DbContext.Where(X => X.ValidFromDate == date);
                    db.DbContext.RemoveRange(data);
                }
                db.SaveChanges();
            }
        }
    }
}
