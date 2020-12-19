using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechnicalTest.BusinessLogic.DataModel;

namespace TechnicalTest.BusinessLogic.Interfaces
{
    public interface iDataAccess
    {
        bool InsertXMLData(List<TransactionInfo> transInfo);
        bool InsertCSVData(List<TransactionInfo> transInfo);
        List<TransactionInfo> GetGridData(string curCode, DateTime fromdate, DateTime todate, string status);
    }
}
