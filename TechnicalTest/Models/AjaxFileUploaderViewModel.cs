using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TechnicalTest.Models
{
    public class AjaxFileUploaderViewModel
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string UploadURL { get; set; }
        public string AllowedFileExtension { get; set; }
        public bool IsReadOnly { get; set; }
        public string FileName { get; set; }
        public string FilePathOrGuid { get; set; }
        public string ExtraClass { get; set; }
    }
}