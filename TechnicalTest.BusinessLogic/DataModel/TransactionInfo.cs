using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace TechnicalTest.BusinessLogic.DataModel
{
    [Serializable]
    [XmlRoot("Transactions"), XmlType("Transactions")]
    public class TransactionInfo
    {
        private List<PaymentDetailInfo> _PaymentDetailInfo;
        public List<PaymentDetailInfo> PaymentDetailInfo
        {
            get { return _PaymentDetailInfo; }
            set { _PaymentDetailInfo = value; }
        }
        public TransactionInfo()
        {
            PaymentDetailInfo = new List<PaymentDetailInfo>();
        }
        public string Id { get; set; }
        public string Status { get; set; }
        public DateTime TransactionDate { get; set; }
    }
}
