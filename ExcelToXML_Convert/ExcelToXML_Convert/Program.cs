using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelToXML_Convert
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileName = args[0];
            var xml_content = openFile(fileName);
            saveXML(xml_content, args[1] );
        }

        static void saveXML(string content, string filename)
        {
            System.IO.File.WriteAllText(filename, content);
        }
        static string openFile(string fileName)
        {
            Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(fileName);
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

            var retVal = @"<?xml version=""1.0"" encoding=""utf - 8""?>\n"  ;
            retVal += @"<XmlOrders XMLns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns: xsd = ""http://www.w3.org/2001/XMLSchema"" > \n";
            retVal += " <XmlOrderInfos>\n";

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;


            for (int i = 2; i <= rowCount; i++)
            {
                retVal += "     <XmlOrderInfo>\n";
                //for (int j = 1; j <= colCount; j++)
                //{                    
                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {
                    retVal += "         <ProductCodes>\n";
                    retVal += "             <Product_Code>" + xlRange.Cells[i, 1].Value2.ToString() + "</Product_Code>\n";
                    retVal += "         </ProductCodes>\n";
                    retVal += "         <Customer_Name>" + xlRange.Cells[i, 3].Value2.ToString() + "</Customer_Name>\n";
                    retVal += "         <Customer_Reference>" + xlRange.Cells[i, 2].Value2.ToString() + "</Customer_Reference>\n";
                    retVal += "         <Address></Address>\n";
                    retVal += "         <Email_Address></Email_Address>\n";
                    retVal += "         <Additional_Email_Address></Additional_Email_Address>\n";
                    retVal += "         <Language_Code></Language_Code>\n";
                    retVal += "         <Additional_Language_Code></Additional_Language_Code>\n";
                    retVal += "         <Dispatch_Method></Dispatch_Method>\n";
                    retVal += "         <Contact_Name></Contact_Name>\n";
                }
                //}
                retVal += "     </XmlOrderInfo>\n";
            }
            retVal += " </XmlOrders>\n";

            //cleanup  
            GC.Collect();
            GC.WaitForPendingFinalizers();
             
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release  
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return retVal;
        }


    }
}
