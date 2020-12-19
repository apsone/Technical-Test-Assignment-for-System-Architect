using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using TechnicalTest.BusinessLogic.Interfaces;
using TechnicalTest.BusinessLogic.BS;
using System.Security.Principal;
using System.Security.Claims;

namespace TechnicalTest.Repositories
{
    public class AppRepositoryManager
    {
        public AppRepositoryManager()
        {

        }
        ~AppRepositoryManager()
        { }
        public static iCSVReader CSVReaderManager
        {
            get
            {
                return new CSVReaderBS();
            }
        }

        public static iXMLReader XMLReaderManager
        {
            get
            {
                return new XMLReaderBS();
            }
        }

        public static iDataAccess DataAccessManager
        {
            get
            {
                return new DataAccessBS();
            }
        }
    }
}