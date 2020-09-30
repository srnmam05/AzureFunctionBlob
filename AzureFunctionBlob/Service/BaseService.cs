using System;
using System.Collections.Generic;
using System.Text;

namespace AzureFunctionBlob.Service
{
    /// <summary>
    /// 連接資料庫用
    /// </summary>
    public class BaseService
    {
        protected DbContext db;
        public BaseService()
        {
            var sql = Environment.GetEnvironmentVariable("");
            var optionsBulider = new DbContextOptionsBuilder<DbContext>();
            optionsBulider.UseSqlServer(sql);
            db = new DbContext(optionsBulider.Options);
        }
    }
}
