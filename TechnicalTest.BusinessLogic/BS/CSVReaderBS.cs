using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Data;
using System.Xml.Linq;
using System.Globalization;
using TechnicalTest.BusinessLogic.DataModel;
using TechnicalTest.BusinessLogic.Interfaces;

namespace TechnicalTest.BusinessLogic.BS
{
    public class CSVReaderBS : iCSVReader
    {
        /// <summary>  
        /// Return list of products from XML.  
        /// </summary>  
        /// <returns>List of products</returns>  
        public List<TransactionInfo> RetrunListOfProducts(string filepath)
        {
            string basicPath = HttpContext.Current.Server.MapPath("~/FileData/");

            string csvDataPath = string.Concat(basicPath, filepath);
            List<TempCSVInfo> tmpList = new List<TempCSVInfo>();
            TempCSVInfo addTemp;
            List<TransactionInfo> results = new List<TransactionInfo>();
            TransactionInfo addObj;
            PaymentDetailInfo addDetail;
            DataSet l_set = new DataSet();
            string csvData = System.IO.File.ReadAllText(csvDataPath);
            //Execute a loop over the rows.
            string distinctInvNo = string.Empty;
            string format = "dd/MM/yyyy HH:mm:ss";
            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    addTemp = new TempCSVInfo();
                    addTemp.Id = row.Split(',')[0].Trim();
                    addTemp.Amount = Convert.ToDecimal(row.Split(',')[1].Trim());
                    addTemp.CurrencyCode = row.Split(',')[2].Trim();
                    addTemp.TransactionDate = DateTime.ParseExact(row.Split(',')[3].Trim(), format, CultureInfo.InvariantCulture);
                    addTemp.Status = row.Split(',')[4].Trim();
                    tmpList.Add(addTemp);
                    distinctInvNo = row.Split(',')[0].Trim();
                }
            }
            var linq_distinct = (from tmp in tmpList
                                 select new { tmp.Id, tmp.Status }).Distinct();
            foreach (var item in linq_distinct)
            {
                addObj = new TransactionInfo();
                addObj.Id = item.Id;
                addObj.Status = item.Status;
                addObj.TransactionDate = tmpList.Where(c => c.Id == item.Id).OrderByDescending(c => c.TransactionDate).FirstOrDefault().TransactionDate;
                addObj.PaymentDetailInfo = new List<PaymentDetailInfo>();
                foreach (var d_item in tmpList.Where(c => c.Id == item.Id))
                {
                    addDetail = new PaymentDetailInfo();
                    addDetail.Amount = d_item.Amount;
                    addDetail.CurrencyCode = d_item.CurrencyCode;
                    addObj.PaymentDetailInfo.Add(addDetail);
                }
                results.Add(addObj);
            }
            return results;
        }
    }
}
