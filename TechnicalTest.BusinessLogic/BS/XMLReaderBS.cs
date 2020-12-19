using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Data;
using System.Xml.Linq;
using TechnicalTest.BusinessLogic.DataModel;
using TechnicalTest.BusinessLogic.Interfaces;

namespace TechnicalTest.BusinessLogic.BS
{
    public class XMLReaderBS : iXMLReader
    {
        public List<TransactionInfo> RetrunListOfProducts(string filepath)
        {
            string xmlData = HttpContext.Current.Server.MapPath(filepath);
            List<TransactionInfo> results = new List<TransactionInfo>();
            TransactionInfo addObj;
            PaymentDetailInfo addDetail;
            DataSet l_set = new DataSet();
            string xmlPath = HttpContext.Current.Server.MapPath("~/FileData/");
            l_set.ReadXml(string.Concat(xmlPath, filepath));
            if (l_set != null)
            {
                DataTable l_tbl = new DataTable();
                DataTable d_tbl = new DataTable();
                l_tbl = l_set.Tables["Transaction"];
                d_tbl = l_set.Tables["PaymentDetails"];
                foreach (DataRow dr in l_tbl.Rows)
                {
                    addObj = new TransactionInfo();
                    addObj.Id = dr["id"].ToString();
                    addObj.Status = dr["Status"].ToString();
                    addObj.TransactionDate = Convert.ToDateTime(dr["TransactionDate"].ToString());
                    DataTable tblFiltered = d_tbl.AsEnumerable()
                          .Where(row => row.Field<Int32>("Transaction_Id") == Convert.ToInt32(dr["Transaction_Id"]))
                          .CopyToDataTable();
                    addObj.PaymentDetailInfo = new List<PaymentDetailInfo>();
                    foreach (DataRow detail in tblFiltered.Rows)
                    {
                        addDetail = new PaymentDetailInfo();
                        addDetail.Amount = Convert.ToDecimal(detail["Amount"].ToString());
                        addDetail.CurrencyCode = detail["CurrencyCode"].ToString();
                        addObj.PaymentDetailInfo.Add(addDetail);
                    }
                    results.Add(addObj);
                }
            }


            return results;
        }

    }
}
