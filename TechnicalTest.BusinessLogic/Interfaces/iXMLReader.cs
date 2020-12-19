using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechnicalTest.BusinessLogic.DataModel;

namespace TechnicalTest.BusinessLogic.Interfaces
{
    public interface iXMLReader
    {
        List<TransactionInfo> RetrunListOfProducts(string filepath);
    }
}
