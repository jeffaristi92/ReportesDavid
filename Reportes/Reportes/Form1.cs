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
            cargarDocumentos();
        }

        void cargarDocumentos()
        {
            string[] strDocumentos = Directory.GetFiles(RutasProcesamiento.Entrada);
            progressBar1.Maximum = strDocumentos.Length;
            progressBar1.Minimum = 0;

            if (strDocumentos.Length > 0)
            {
                string[] strArchivo1 = Directory.GetFiles(RutasProcesamiento.Entrada, "P1_*");
                string[] strArchivo2 = Directory.GetFiles(RutasProcesamiento.Entrada, "P2_*");
                string[] strArchivo3 = Directory.GetFiles(RutasProcesamiento.Entrada, "P3_*");

                if (strArchivo1.Length > 0 && strArchivo2.Length > 0 && strArchivo3.Length > 0)
                {
                    string strReporte = procesarDocumento(strArchivo1[0], strArchivo2[0], strArchivo3[0]);

                    if (!strReporte.Equals(string.Empty)) {
                        string strRutaSalida = Path.Combine(RutasProcesamiento.Salida,string.Format("ReporteRiplay_{0:yyyyMMddHHmmss}.txt",DateTime.Now));
                        File.WriteAllText(strRutaSalida,strReporte);
                    }
                }
            }
        }

        string procesarDocumento(string strArchivo1, string strArchivo2, string strArchivo3)
        {
            DataTable dtPrimerArchivo = cargarPrimerArchivo(strArchivo1);
            DataTable dtSegundoArchivo = cargarSegundoArchivo(strArchivo2);
            DataTable dtTercerArchivo = cargarTercerArchivo(strArchivo3);

            List<string> lstArchivosNoProcesados = getArchivosNoProcesados(dtPrimerArchivo, dtSegundoArchivo);
            List<string[]> lstArchivosNoCargados = getArchivosNoCargados(dtSegundoArchivo, dtTercerArchivo);

            return generarReporte(lstArchivosNoProcesados, lstArchivosNoCargados).ToString();
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
                        sbReporte.Append(archivo);
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
            string strRutaArchivo = Path.Combine(RutasProcesamiento.Adicionales[0], archivo);

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

        List<string[]> getArchivosNoCargados(DataTable dtSegundoArchivo, DataTable dtTercerArchivo)
        {
            List<string[]> lstArchivosNoCargados = new List<string[]>();
            List<string> lstArchivosCargados = getArchivosCargados(dtTercerArchivo);

            foreach (DataRow fila in dtSegundoArchivo.Rows)
            {
                string strNroDoc = getNroDocumento(fila[0].ToString());
                if (!lstArchivosCargados.Contains(strNroDoc))
                {
                    string strMotivo = getMotivo(fila[0].ToString());

                    lstArchivosNoCargados.Add(new string[] { fila[0].ToString(), strMotivo });
                }
            }

            return lstArchivosNoCargados;
        }

        List<string> getArchivosCargados(DataTable dtTercerArchivo)
        {
            List<string> lstArchivosCargados = new List<string>();

            foreach (DataRow fila in dtTercerArchivo.Rows)
            {
                string doc = fila[8].ToString().Trim();
                string nro = doc.Substring(doc.Length - 8, 7);
                lstArchivosCargados.Add(nro);
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
