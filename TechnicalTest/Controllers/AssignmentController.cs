using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TechnicalTest.Models;
using TechnicalTest.Repositories;
using TechnicalTest.BusinessLogic;
using TechnicalTest.BusinessLogic.DataModel;
using System.Web.Script.Serialization;

namespace TechnicalTest.Controllers
{
    public class AssignmentController : Controller
    {
        [HttpPost]
        public ActionResult UploadFiles()
        {

            string path = Server.MapPath("~/FileData/");
            HttpFileCollectionBase files = Request.Files;
            string physicalpath = string.Empty;
            for (int i = 0; i < files.Count; i++)
            {
                HttpPostedFileBase file = files[i];
                if (file.FileName.EndsWith(".xml"))
                {
                    file.SaveAs(path + file.FileName);
                    physicalpath = file.FileName;
                }
                else
                {
                    throw new Exception("Invalid File Type. File type must be XML");
                }
            }
            return Json(physicalpath);
        }
        [HttpPost]
        public ActionResult UploadCSVFiles()
        {

            string path = Server.MapPath("~/FileData/");
            HttpFileCollectionBase files = Request.Files;
            string physicalpath = string.Empty;
            for (int i = 0; i < files.Count; i++)
            {
                HttpPostedFileBase file = files[i];
                if (file.FileName.EndsWith(".csv"))
                {
                    file.SaveAs(path + file.FileName);
                    physicalpath = file.FileName;
                }
                else
                {
                    throw new Exception("Invalid File Type. File type must be XML");
                }
            }
            return Json(physicalpath);
        }
        public ActionResult ViewUploadedResults()
        {
            return View();
        }
        public ActionResult CSVAssignment()
        {
            FileUploadViewModel viewModel = new FileUploadViewModel();
            return View(viewModel);
        }
        public ActionResult XMLAssignment()
        {
            FileUploadViewModel viewModel = new FileUploadViewModel();
            return View(viewModel);
        }
        [HttpPost, ValidateAntiForgeryToken]
        public ActionResult XMLAssignment(FileUploadViewModel paramModel)
        {
            if (ModelState.IsValid)
            {
                if (paramModel != null)
                {
                    List<TransactionInfo> results = AppRepositoryManager.XMLReaderManager.RetrunListOfProducts(paramModel.XMLFileLocation);
                    bool IsSuccessDB = AppRepositoryManager.DataAccessManager.InsertXMLData(results);
                    ViewBag.Result = "true";
                    return View(paramModel);
                }
                else
                {
                    ViewBag.Result = "false";
                    FileUploadViewModel newModel = new FileUploadViewModel();
                    return View(newModel);
                }
            }
            FileUploadViewModel createmodel = new FileUploadViewModel();
            return View(createmodel);
        }
        [HttpPost, ValidateAntiForgeryToken]
        public ActionResult CSVAssignment(FileUploadViewModel paramModel)
        {
            if (ModelState.IsValid)
            {
                if (paramModel != null)
                {
                    List<TransactionInfo> results = AppRepositoryManager.CSVReaderManager.RetrunListOfProducts(paramModel.CSVFileLocation);
                    bool IsSuccessDB = AppRepositoryManager.DataAccessManager.InsertCSVData(results);
                    ViewBag.Result = "true";
                    return View(paramModel);
                }
                else
                {
                    ViewBag.Result = "false";
                    FileUploadViewModel newModel = new FileUploadViewModel();
                    return View(newModel);
                }
            }
            FileUploadViewModel createmodel = new FileUploadViewModel();
            return View(createmodel);
        }
        [HttpPost]
        public ActionResult GetUploadedData(string curCode, string status,string fromdate, string todate)
        {
            List<TransactionInfo> lstFinalResult = new List<TransactionInfo>();
            DateTime fromDate;
            DateTime toDate;
            lstFinalResult = AppRepositoryManager.DataAccessManager.GetGridData(curCode, DateTime.Now, DateTime.Now, status);
            if (fromdate != string.Empty && todate != string.Empty)
            {
                fromDate = StringToDateTime_MMddyyyy(fromdate);
                toDate = StringToDateTime_MMddyyyy(todate);

                var date_linq_filter = from dateResult in lstFinalResult
                                       where dateResult.TransactionDate.Date >= fromDate.Date && dateResult.TransactionDate.Date <= toDate.Date
                                       select dateResult;
                if (date_linq_filter.Count() > 0)
                {
                    lstFinalResult = new List<TransactionInfo>();
                    lstFinalResult = date_linq_filter.ToList<TransactionInfo>();
                }
                else
                {
                    lstFinalResult = new List<TransactionInfo>();
                }
            }

            JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
            string result = javaScriptSerializer.Serialize(lstFinalResult.ToList());
            return Json(result, JsonRequestBehavior.AllowGet);
        }

        private DateTime StringToDateTime_MMddyyyy(String txtDate)
        {
            try
            {
                string Vdatetime;
                string Vday;
                string Vmonth;
                string Vyear;
                Vdatetime = txtDate.Trim();
                Vday = Vdatetime.Substring(0, 2);
                Vmonth = Vdatetime.Substring(3, 2);
                Vyear = Vdatetime.Substring(6, 4);
                return new DateTime(Convert.ToInt16(Vyear), Convert.ToInt16(Vmonth), Convert.ToInt16(Vday), DateTime.Now.Hour,
                                     DateTime.Now.Minute, DateTime.Now.Second);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }
    }
}