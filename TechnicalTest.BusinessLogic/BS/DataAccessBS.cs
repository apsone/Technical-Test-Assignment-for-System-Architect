using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechnicalTest.BusinessLogic.DataModel;
using TechnicalTest.BusinessLogic.Interfaces;

namespace TechnicalTest.BusinessLogic.BS
{
    public class DataAccessBS : iDataAccess
    {
        public bool InsertXMLData(List<TransactionInfo> transInfo)
        {
            bool result = false;
            try
            {
                tblExcel _xml;
                tblXMLPaymentDetail _details;

                using (Entities db = new Entities())
                {
                    foreach (var item in transInfo)
                    {
                        _xml = new tblExcel();
                        _xml.Syskey = Guid.NewGuid();
                        _xml.TransID = item.Id;
                        _xml.TransDate = item.TransactionDate;
                        _xml.RecordStatus = item.Status;
                        db.tblExcels.Add(_xml);
                        foreach (var det in item.PaymentDetailInfo)
                        {
                            _details = new tblXMLPaymentDetail();
                            _details.Syskey = Guid.NewGuid();
                            _details.XMLId = _xml.Syskey;
                            _details.Amount = det.Amount;
                            _details.CurrencyCode = det.CurrencyCode;
                            db.tblXMLPaymentDetails.Add(_details);
                        }
                    }

                    db.SaveChanges();
                    result = true;
                }
            }
            catch (Exception ex)
            {
                result = false;
            }
            return result;
        }
        public bool InsertCSVData(List<TransactionInfo> transInfo)
        {
            bool result = false;
            try
            {
                tblCSV _xml;
                tblCSVPaymentDetail _details;

                using (Entities db = new Entities())
                {
                    foreach (var item in transInfo)
                    {
                        _xml = new tblCSV();
                        _xml.Syskey = Guid.NewGuid();
                        _xml.TransID = item.Id;
                        _xml.TransDate = item.TransactionDate;
                        _xml.RecordStatus = item.Status;
                        db.tblCSVs.Add(_xml);
                        foreach (var det in item.PaymentDetailInfo)
                        {
                            _details = new tblCSVPaymentDetail();
                            _details.Syskey = Guid.NewGuid();
                            _details.CSVId = _xml.Syskey;
                            _details.Amount = det.Amount;
                            _details.CurrencyCode = det.CurrencyCode;
                            db.tblCSVPaymentDetails.Add(_details);
                        }
                    }

                    db.SaveChanges();
                    result = true;
                }
            }
            catch (Exception ex)
            {
                result = false;
            }
            return result;
        }
        public List<TransactionInfo> GetGridData(string curCode, DateTime fromdate, DateTime todate, string status)
        {
            List<TransactionInfo> results = new List<TransactionInfo>();
            TransactionInfo addObj;
            PaymentDetailInfo addDetail;
            using (Entities entity = new Entities())
            {

                var xml_queriedData = from x in entity.tblExcels.AsEnumerable()
                                      join d in entity.tblXMLPaymentDetails.AsEnumerable() on x.Syskey equals d.XMLId
                                      where d.CurrencyCode == (curCode != string.Empty ? curCode : d.CurrencyCode)
                                      //&& (x.TransDate >= fromdate && x.TransDate <= todate)
                                      && x.RecordStatus == (status != string.Empty ? status : x.RecordStatus)
                                      select new { x, d };
                var xml_distinctData = (from d in xml_queriedData
                                        select new { d.x }).Distinct();
                foreach (var item in xml_distinctData)
                {
                    addObj = new TransactionInfo();
                    addObj.Id = item.x.TransID;
                    addObj.TransactionDate = item.x.TransDate;
                    addObj.Status = item.x.RecordStatus;
                    addObj.PaymentDetailInfo = new List<PaymentDetailInfo>();
                    foreach (var d_item in xml_queriedData.Where(c => c.d.XMLId == item.x.Syskey))
                    {
                        addDetail = new PaymentDetailInfo();
                        addDetail.Amount = d_item.d.Amount;
                        addDetail.CurrencyCode = d_item.d.CurrencyCode;
                        addObj.PaymentDetailInfo.Add(addDetail);
                    }
                    results.Add(addObj);
                }


                var csv_queriedData = from x in entity.tblCSVs.AsEnumerable()
                                      join d in entity.tblCSVPaymentDetails.AsEnumerable() on x.Syskey equals d.CSVId
                                      where d.CurrencyCode == (curCode != string.Empty ? curCode : d.CurrencyCode)
                                      //&& (x.TransDate >= fromdate && x.TransDate <= todate)
                                      && x.RecordStatus == (status != string.Empty ? status : x.RecordStatus)
                                      select new { x, d };
                var csv_distinctData = (from d in csv_queriedData
                                        select new { d.x }).Distinct();
                foreach (var item in csv_distinctData)
                {
                    addObj = new TransactionInfo();
                    addObj.Id = item.x.TransID;
                    addObj.TransactionDate = item.x.TransDate;
                    addObj.Status = item.x.RecordStatus;
                    addObj.PaymentDetailInfo = new List<PaymentDetailInfo>();
                    foreach (var d_item in csv_queriedData.Where(c => c.d.CSVId == item.x.Syskey))
                    {
                        addDetail = new PaymentDetailInfo();
                        addDetail.Amount = d_item.d.Amount;
                        addDetail.CurrencyCode = d_item.d.CurrencyCode;
                        addObj.PaymentDetailInfo.Add(addDetail);
                    }
                    results.Add(addObj);
                }
            }
            return results;
        }
    }
}
