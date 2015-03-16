using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using System.Data.OleDb;

namespace ProcesamientoXML
{
    public partial class Form1 : Form
    {
        DataSet dsReferencias = new DataSet("Referencias");
        Rutas RutasProcesamiento;
        public Form1()
        {
            InitializeComponent();
            Text = "Reportes";
            startProcess();
        }

        public void startProcess()
        {
            Show();
            cargarRutas();
            string rutaReferencia = Path.Combine(RutasProcesamiento.Referencias, "Motivos.txt");
            dsReferencias.Tables.Add(cargarReferencia(rutaReferencia, "Motivos", "|", 2, new string[] { "LLAVE", "VALOR" }));
            try
            {
                List<clsReporte> lstReportes = new List<clsReporte>();

                lstReportes.Add(cargarDocumentos(RutasProcesamiento.Entrada, 1, "ORDERS"));
                if (RutasProcesamiento.Adicionales.Count >= 2)
                {
                    lstReportes.Add(cargarDocumentos(RutasProcesamiento.Adicionales[1], 2, "INVRPT"));
                }
                if (RutasProcesamiento.Adicionales.Count >= 3)
                {
                    lstReportes.Add(cargarDocumentos(RutasProcesamiento.Adicionales[2], 2, "SLSRPT"));
                }
                generarResumen(lstReportes);

            }
            catch (Exception exc)
            {
                MessageBox.Show(string.Format("[error] {0}", exc.Message));
            }
        }

        void generarResumen(List<clsReporte> lstReportes) {
            StringBuilder sbResumen = new StringBuilder();
            sbResumen.AppendLine(DateTime.Now.ToString("dd/MM/yyyy"));
            sbResumen.AppendLine("DOCUMENTO,GENERADOS,BACKUP,NO PROCESADOS,CEN,ERROR");
            foreach(clsReporte reporte in lstReportes){
                sbResumen.AppendLine(reporte.ToString());   
            }
            string strRutaSalida = Path.Combine(RutasProcesamiento.Salida, string.Format("Resumen_ReporteRiplay_{0:yyyyMMddHHmmss}.txt", DateTime.Now));
            File.WriteAllText(strRutaSalida, sbResumen.ToString());
        }

        clsReporte cargarDocumentos(string strRuta, int tipo, string doc)
        {
            clsReporte objReporte = new clsReporte();
            string[] strDocumentos = Directory.GetFiles(strRuta);

            if (strDocumentos.Length > 0)
            {
                string[] strArchivo1 = Directory.GetFiles(strRuta, "P1_*");
                string[] strArchivo2 = Directory.GetFiles(strRuta, "P2_*");
                string[] strArchivo3 = Directory.GetFiles(strRuta, "P3_*");

                if (strArchivo1.Length > 0 && strArchivo2.Length > 0 && strArchivo3.Length > 0)
                {
                    objReporte = procesarDocumento(strArchivo1[0], strArchivo2[0], strArchivo3[0], tipo,doc);

                    if (!objReporte.Reporte.Equals(string.Empty))
                    {
                        string strRutaSalida = Path.Combine(RutasProcesamiento.Salida, string.Format("{0}_ReporteRiplay_{1:yyyyMMddHHmmss}.txt", doc, DateTime.Now));
                        File.WriteAllText(strRutaSalida, objReporte.Reporte);
                    }
                }
            }
            return objReporte;
        }

        clsReporte procesarDocumento(string strArchivo1, string strArchivo2, string strArchivo3, int tipo,string doc)
        {
            clsReporte objReporte = new clsReporte();

            DataTable dtPrimerArchivo = cargarPrimerArchivo(strArchivo1);
            DataTable dtSegundoArchivo = cargarSegundoArchivo(strArchivo2);
            DataTable dtTercerArchivo = cargarTercerArchivo(strArchivo3);

            List<string> lstArchivosNoProcesados = getArchivosNoProcesados(dtPrimerArchivo, dtSegundoArchivo);
            List<string[]> lstArchivosNoCargados = getArchivosNoCargados(dtSegundoArchivo, dtTercerArchivo, tipo);

            objReporte.Documento = doc;
            objReporte.Generados = dtPrimerArchivo.Rows.Count;
            objReporte.Backup = dtSegundoArchivo.Rows.Count;
            objReporte.NoProcesados = lstArchivosNoProcesados.Count;
            objReporte.CEN = dtTercerArchivo.Rows.Count;
            objReporte.Error = lstArchivosNoCargados.Count;
            objReporte.Reporte = generarReporte(lstArchivosNoProcesados, lstArchivosNoCargados).ToString();

            return objReporte;
        }

        StringBuilder generarReporte(List<string> lstArchivosNoProcesados, List<string[]> lstArchivosNoCargados)
        {
            StringBuilder sbReporte = new StringBuilder();

            if (lstArchivosNoProcesados.Count > 0 || lstArchivosNoCargados.Count > 0)
            {
                sbReporte.AppendLine(string.Format("REPORTE REPLAY {0}", DateTime.Now.ToString("dd/MM/yyyy")));
                if (lstArchivosNoProcesados.Count > 0)
                {
                    sbReporte.AppendLine("NO PROCESADOS");
                    sbReporte.AppendLine("-------------------------------");
                    foreach (string archivo in lstArchivosNoProcesados)
                    {
                        sbReporte.AppendLine(archivo);
                    }
                    sbReporte.AppendLine("-------------------------------");
                }

                if (lstArchivosNoCargados.Count > 0)
                {
                    sbReporte.AppendLine("NO CARGADOS AL CEN");
                    sbReporte.AppendLine("-------------------------------");
                    foreach (string[] archivo in lstArchivosNoCargados)
                    {
                        sbReporte.AppendLine(string.Format("{0} - {1}", archivo[0], archivo[1]));
                    }
                    sbReporte.AppendLine("-------------------------------");
                }
            }

            return sbReporte;
        }

        string getMotivo(string archivo)
        {
            string strMotivo = string.Empty;
            string strAux = RutasProcesamiento.Adicionales[0];
            string strRutaArchivo = Path.Combine(strAux, archivo);

            if (File.Exists(strRutaArchivo))
            {
                string[] contenidoArchivo = File.ReadAllLines(strRutaArchivo);
                string strBusqueda = string.Empty;
                if (contenidoArchivo.Length > 0)
                {
                    string[] linea = contenidoArchivo[0].Split('|');
                    strBusqueda = linea[1];
                }
                string strConsulta = string.Format("LLAVE = '{0}'", strBusqueda);
                DataRow[] resultado = dsReferencias.Tables["Motivos"].Select(strConsulta);

                if (resultado.Length > 0)
                {
                    strMotivo = resultado[0][1].ToString();
                }
            }

            return strMotivo;
        }

        List<string> getArchivosNoProcesados(DataTable dtPrimerArchivo, DataTable dtSegundoArchivo)
        {
            List<string> lstArchivosNoProcesados = new List<string>();

            foreach (DataRow fila in dtPrimerArchivo.Rows)
            {
                string strConsulta = string.Format("Archivo like '%_{0}_%'", getNroDocumento(fila[0].ToString()));
                DataRow[] resultado = dtSegundoArchivo.Select(strConsulta);
                if (resultado.Length == 0)
                {
                    lstArchivosNoProcesados.Add(fila[0].ToString());
                }
            }

            return lstArchivosNoProcesados;
        }

        List<string[]> getArchivosNoCargados(DataTable dtSegundoArchivo, DataTable dtTercerArchivo, int tipo)
        {
            List<string[]> lstArchivosNoCargados = new List<string[]>();
            List<string> lstArchivosCargados = getArchivosCargados(dtTercerArchivo, tipo);

            foreach (DataRow fila in dtSegundoArchivo.Rows)
            {
                string strValor = string.Empty;
                switch(tipo){
                    case 1:
                        strValor = getNroDocumento(fila[0].ToString()); 
                        break;
                    case 2:
                        strValor = getSNRF(fila[0].ToString());
                        break;
                    default:
                        break;
                }
                if (!lstArchivosCargados.Contains(strValor))
                {
                    string strMotivo = getMotivo(fila[0].ToString());
                    lstArchivosNoCargados.Add(new string[] { fila[0].ToString(), strMotivo });
                }
            }

            return lstArchivosNoCargados;
        }

        List<string> getArchivosCargados(DataTable dtTercerArchivo, int tipo)
        {
            List<string> lstArchivosCargados = new List<string>();

            foreach (DataRow fila in dtTercerArchivo.Rows)
            {
                string doc = string.Empty;
                switch (tipo)
                {
                    case 1:
                        doc = fila[8].ToString().Trim();
                        if (doc.Length > 7)
                        {
                            doc = doc.Substring(doc.Length - 8, 7);
                        }
                        else
                        {
                            doc = string.Empty;
                        }
                        break;
                    case 2:
                        doc = fila[7].ToString().Trim();
                        break;
                    default:
                        break;
                }

                if (!doc.Equals(string.Empty))
                {
                    lstArchivosCargados.Add(doc);
                }
            }

            return lstArchivosCargados;
        }

        string getNroDocumento(string nombreDocumento)
        {
            string strNroDocumento = string.Empty;
            string[] lstPartes = nombreDocumento.Split('_');

            if (lstPartes.Length > 1)
            {
                strNroDocumento = lstPartes[1];
            }

            return strNroDocumento;
        }

        string getSNRF(string nombreDocumento)
        {
            string strNroDocumento = string.Empty;
            string[] lstPartes = Path.GetFileNameWithoutExtension(nombreDocumento).Split('_');

            if (lstPartes.Length > 3)
            {


                for (int i = 3; i < lstPartes.Length; i++)
                {
                    if (i == 3)
                    {
                        strNroDocumento = lstPartes[i];
                    }
                    else
                    {
                        strNroDocumento = string.Format("{0}_{1}", strNroDocumento, lstPartes[3]);
                    }
                }
            }

            return strNroDocumento;
        }

        DataTable cargarPrimerArchivo(string strDocumento)
        {
            DataTable dtPrimerArchivo = ExcelToDataTable(strDocumento);
            return dtPrimerArchivo;
        }

        DataTable cargarSegundoArchivo(string strDocumento)
        {
            DataTable dtSegundoArchivo = new DataTable();
            DataColumn columna1 = new DataColumn("Archivo");
            dtSegundoArchivo.Columns.Add(columna1);

            string[] strLineas = File.ReadAllLines(strDocumento);

            if (strLineas.Length > 1)
            {
                for (int i = 1; i < strLineas.Length; i++)
                {
                    dtSegundoArchivo.Rows.Add(strLineas[i]);
                }
            }


            return dtSegundoArchivo;
        }

        DataTable cargarTercerArchivo(string strDocumento)
        {
            DataTable dtTercerArchivo = ExcelToDataTable(strDocumento);
            return dtTercerArchivo;
        }

        DataTable ExcelToDataTable(string path)
        {
            var pck = new OfficeOpenXml.ExcelPackage();
            pck.Load(File.OpenRead(path));
            var ws = pck.Workbook.Worksheets.First();
            DataTable tbl = new DataTable();
            bool hasHeader = true;
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
            {
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            var startRow = hasHeader ? 2 : 1;
            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                var row = tbl.NewRow();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
                tbl.Rows.Add(row);
            }
            pck.Dispose();
            return tbl;
        }

        DataTable cargarReferencia(string archivo, string nombre, string separador, int nroColumnas, string[] nombresColumnas)
        {
            DataTable tabla = new DataTable(nombre);
            for (int i = 0; i < nroColumnas; i++)
            {
                tabla.Columns.Add(nombresColumnas[i]);
            }

            string[] filas = System.IO.File.ReadAllLines(archivo);

            foreach (string fila in filas)
            {
                string[] campos = fila.Split(separador[0]);
                tabla.Rows.Add(campos);
            }

            return tabla;
        }

        void cargarRutas()
        {
            RutasProcesamiento = new Rutas();
            RutasProcesamiento.Archivo = Path.Combine(Directory.GetCurrentDirectory(), "rutas.txt");
            RutasProcesamiento.cargarRutas();
        }
    }
}
