using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PDF_Process;
using System.Configuration;
using System.Data;
using Microsoft.Graph;

namespace MH_INVOICE_V1
{
    class Program
    {
        static void Main(string[] args)
        {
            string pdfPath = ConfigurationManager.AppSettings["pdfPath"];
            string opPath = ConfigurationManager.AppSettings["opPath"];
            string cpPath = ConfigurationManager.AppSettings["cpPath"];
            string filePath = ConfigurationManager.AppSettings["filePath"];
            

            //Extract PDF data using Tessaract OCR.
            PDF_Process.ExternalMethods externalMethods = new PDF_Process.ExternalMethods();
            externalMethods.ExtractPDFContent(pdfPath);

            //Extract PDF Split range
            externalMethods.ExtractPDFValues();

            PDF_Process.ExternalMethods extractData = new PDF_Process.ExternalMethods();
            DataTable dt = extractData.ExtractPDFValues();
            //Split and Compress PDF
            externalMethods.ExtractRange_Split_Compress(dt, pdfPath, opPath, cpPath);

            //one drive upload
            One_Drive_Process.ExternalAPIMethods externalApiMethods = new One_Drive_Process.ExternalAPIMethods();
            externalApiMethods.UploadOneDriveFile(filePath);

        }

        
    }
}