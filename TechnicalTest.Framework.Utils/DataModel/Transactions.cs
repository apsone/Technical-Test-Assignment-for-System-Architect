using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace TechnicalTest.Framework.Utils.DataModel
{
    public class Transactions
    {
        [Serializable]
        [XmlRoot("Transactions"), XmlType("Transactions")]
        public class Transaction
        {            
            private PaymentDetails _PaymentDetailInfo;
            public PaymentDetails PaymentDetailInfo
            {
                get { return _PaymentDetailInfo; }
                set { _PaymentDetailInfo = value; }
            }
            public Transaction()
            {
                PaymentDetailInfo = new PaymentDetails();
            }
            public string Id { get; set; }
            public string Status { get; set; }
        }
    }
}
