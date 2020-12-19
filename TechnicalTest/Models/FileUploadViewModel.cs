using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace TechnicalTest.Models
{
    public class FileUploadViewModel
    {
        [Required(ErrorMessage = "Required")]
        [Display(Name = "Upload XML File")]
        public string XMLFileLocation { get; set; }

        [Required(ErrorMessage = "Required")]
        [Display(Name = "Upload CSV File")]
        public string CSVFileLocation { get; set; }
    }
}