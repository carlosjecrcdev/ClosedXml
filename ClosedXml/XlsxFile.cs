using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;

namespace ClosedXml
{
    internal class XlsxFile
    {
        public static void GenerarArchivoExcel<T>(List<T> listado)
        {
            var ruta = ConfigurationManager.AppSettings["rutaExcel"];

            if (ruta == null)
            {
                throw new Exception("No se encontró ruta para guardar el fichero.");
            }

            using(var workbook = new XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add("Informe");
                sheet.Cell(1, 1).InsertData(listado);
                workbook.SaveAs(ruta);
            }

        }
    }
}
