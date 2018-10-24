using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Drawing.Printing;
using PdfiumViewer;
using System.Linq;
using System.Management;

namespace Vls.Printer
{
    public class SupportedPrinters
    {
        private static readonly string[] _PRINTERS = {
            //"canon lbp2900",
            //"canon lbp3500",
            //"fax"
            "abc",
        };
        private static readonly HashSet<string> names = new HashSet<string>(_PRINTERS);

        public static bool support(string pname)
        {
            if (pname == null)
            {
                return false;
            }
            return !names.Contains(pname.ToLower());
        }
    }

    public class Result
    {
        public bool success { get; set; }
        public string message { get; set; }
        public Result(bool Success, string Message)
        {
            success = Success;
            message = Message;
        }
        public Result() : this(false, "") {}
    }

    public class Printing
    {
        Result result = new Result();
        public async Task<object> Invoke(dynamic input)
        {
            return this.Print((string)input.pathName, (string)input.printerName, (string)input.paperSize, (Int16)input.copies);
        }

        public object Print(string filePath, string printerName, string paperSize, Int16 copies)
        {   
            if (!SupportedPrinters.support(printerName))
            {
                result.success = false;
                result.message = String.Format("Không hỗ trợ máy in {0}", printerName);
                return result;
            }
            return Helper.Printing(filePath, printerName, paperSize, copies);
        }


        static class Helper
        {
            public static object Printing(string filePath, string printerName, string paperSize, Int16 copies)
            {
                Result result = new Result();
                using (PdfDocument _document = PdfDocument.Load(filePath))
                {

                    using (var document = _document.CreatePrintDocument(PdfPrintMode.ShrinkToMargin))
                    {
                        document.PrinterSettings.PrinterName = printerName;

                        IEnumerable<PaperSize> paperSizes = document.PrinterSettings.PaperSizes.Cast<PaperSize>(); // Các loại khổ giấy được hỗ trợ bởi drive máy in
                        PaperSize paperSz = paperSizes.FirstOrDefault<PaperSize>(size => size.PaperName.ToLower() == paperSize.ToLower());
                        if (paperSz == null)
                        {
                            result.message = String.Format("Không hỗ trợ khổ giấy {0}", paperSize);
                        }
                        else
                        {
                            document.PrinterSettings.Copies = copies;
                            document.DefaultPageSettings.PaperSize = paperSz;
                            document.DefaultPageSettings.Landscape = true;
                            try
                            {
                            document.Print();
                            result.success = true;
                            result.message = "In thành công";

                            }
                            catch (Exception e)
                            {
                                result.message = String.Format("Có lỗi xẩy ra. {0}", e.Message);
                            }
                        }

                    }
                }
                return result;
            }
        }
    }

    public class GetPrinterStatus
    {
        public async Task<object> Invoke(dynamic input)
        {
            return this.GetPrinters();
        }

        public object GetPrinters()
        {
            return Helper.GetPrinterStatus();
        }


        static class Helper
        {
            public static object GetPrinterStatus()
            {
                var result = new Dictionary<string, bool>();
                ManagementScope scope = new ManagementScope(@"\root\cimv2");
                scope.Connect();

                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");

                var allPrinter = searcher.Get();
                foreach (ManagementObject printer in allPrinter)
                {
                    var printerName = printer["Name"].ToString() ;
                    if (SupportedPrinters.support(printerName))
                    {
                        result[printerName] = !printer["WorkOffline"].ToString().ToLower().Equals("true");
                    }
                }
                return result;
            }
        }
    }
}
