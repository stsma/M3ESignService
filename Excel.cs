using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel.Application;

namespace M3ESignService
{
    public class ExcelService
    {
        private readonly Excel oExcel = new Excel();
        public CodeResult Export()
        {
            var cr = CodeResult.NoData;
            //add workbook to excel
            var oWorkBook = oExcel.Workbooks.Add();

            try
            {

            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
