using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Reportes
{
    class clsReporte
    {
        public string Documento { get; set; }
        public int Generados { get; set; }
        public int Backup { get; set; }
        public int NoProcesados { get; set; }
        public int CEN { get; set; }
        public int Error { get; set; }
        public string Reporte { get; set; }

        public clsReporte() {
            Documento = string.Empty;
            Generados = 0;
            Backup = 0;
            NoProcesados = 0;
            CEN = 0;
            Error = 0;
            Reporte = string.Empty;
        }

        public string ToString() {
            return string.Format("{0},{1},{2},{3},{4},{5}",Documento,Generados,Backup,NoProcesados,CEN,Error);
        }
    }
}
