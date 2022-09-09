using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXml
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<Estudiante> lista = new List<Estudiante>();

            Estudiante primerEstudiante = new Estudiante();
            primerEstudiante.matricula = 919154;
            primerEstudiante.nombre = "coder";
            primerEstudiante.apellidoPaterno = "anonimo";
            primerEstudiante.apellidoMaterno = "anonimoMaterno";
            primerEstudiante.grado = 1;
            lista.Add(primerEstudiante);

            Estudiante segundoEstudiante = new Estudiante();
            segundoEstudiante.matricula = 919154;
            segundoEstudiante.nombre = "coder";
            segundoEstudiante.apellidoPaterno = "anonimo";
            segundoEstudiante.apellidoMaterno = "anonimoMaterno";
            segundoEstudiante.grado = 1;
            lista.Add(segundoEstudiante);


            XlsxFile.GenerarArchivoExcel<Estudiante>(lista);
        }
    }
}
