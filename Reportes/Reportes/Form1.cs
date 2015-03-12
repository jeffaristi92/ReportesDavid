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

        public void startProcess() {
            Show();
            cargarRutas();
            string rutaReferencia = Path.Combine(RutasProcesamiento.Referencias, "configuraciones.txt");
            dsReferencias.Tables.Add(cargarReferencia(rutaReferencia, "Configuraciones", "=", 2, new string[] { "LLAVE", "VALOR" }));
            cargarDocumentos();
        }

        void cargarDocumentos()
        {
            string[] strDocumentos = Directory.GetFiles(RutasProcesamiento.Entrada);
            progressBar1.Maximum = strDocumentos.Length;
            progressBar1.Minimum = 0;

            foreach (string rutaArchivo in strDocumentos)
            {
                try
                {
                    XmlDocument xmlDocumento = new XmlDocument();
                    xmlDocumento.Load(rutaArchivo);
                    xmlDocumento.Prefix = "cbc";

                    string raiz = xmlDocumento.DocumentElement.Name;

                    if (raiz.Equals("Invoice"))
                    {
                        XmlNodeList nodo = xmlDocumento.GetElementsByTagName("cbc:InvoiceTypeCode");

                        if (nodo.Count > 0)
                        {
                            XmlNode nuevoNodo = xmlDocumento.CreateElement("cbc:Note");

                            if (nodo[0].InnerText.Equals("01"))
                            {
                                nuevoNodo.InnerText = dsReferencias.Tables["Configuraciones"].Select(string.Format("LLAVE = 'Invoice01'"))[0][1].ToString();
                                xmlDocumento.DocumentElement.InsertAfter(nuevoNodo, nodo[0]);
                            }
                            else if (nodo[0].InnerText.Equals("03"))
                            {
                                nuevoNodo.InnerText = dsReferencias.Tables["Configuraciones"].Select(string.Format("LLAVE = 'Invoice03'"))[0][1].ToString();
                                xmlDocumento.DocumentElement.InsertAfter(nuevoNodo, nodo[0]);
                            }
                        }
                    }
                    else if (raiz.Equals("DebitNote") || raiz.Equals("CreditNote"))
                    {
                        XmlNode nodoReferencia = xmlDocumento.GetElementsByTagName("cbc:IssueDate")[0];
                        XmlNode nuevoNodo = xmlDocumento.CreateElement("Note", xmlDocumento.NamespaceURI);
                        nuevoNodo.Prefix = "cbc";
                        nuevoNodo.InnerText = dsReferencias.Tables["Configuraciones"].Select(string.Format("LLAVE = '{0}'", raiz))[0][1].ToString();
                        xmlDocumento.DocumentElement.InsertAfter(nuevoNodo, nodoReferencia);
                    }

                    File.Move(rutaArchivo, Path.Combine(RutasProcesamiento.Backup, Path.GetFileName(rutaArchivo)));
                    xmlDocumento.Save(Path.Combine(RutasProcesamiento.Salida, Path.GetFileName(rutaArchivo)));
                }
                catch (Exception exc)
                {
                }
                progressBar1.Value++;
                Refresh();
            }
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
