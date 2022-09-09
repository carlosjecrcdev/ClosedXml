using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXml
{
    internal class XlsxFile
    {
        public static void GenerarArchivoExcel()
        {
            var ruta = ConfigurationManager.AppSettings["rutaExcel"];

            if(ruta== null)
            {
                throw new Exception("No se encontró ruta para guardar el fichero.");
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Mi primer hoja");
                worksheet.Cells("A1").Value = "Hello CoderAnonimo";
                workbook.SaveAs(ruta);
            }
        }
    }
}
