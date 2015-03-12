using System.Data;
using System;
using System.IO;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.Text;
namespace SI.CO.NOVARTIS.DESADV.TXT.EDI
{
    class clsUtil
    {
        public static bool mtdIsNumber(string cadena)
        {
            double aux = 0;
            return double.TryParse(cadena, out aux);
        }

        public static string mtdRemoverCaracteresEspeciales(string cadena, string caracteres)
        {
            if (cadena == null || caracteres == null) return cadena;

            foreach (char caracter in caracteres)
            {
                cadena = cadena.Replace(caracter, ' ');
            }

            return cadena.Trim();
        }

        public static string mtdSelectFromReference(DataSet References, int tabla, string columnaBusqueda, string valorBusqueda, int columnaResultado)
        {
            try
            {
                DataRow[] foundRows = References.Tables[tabla].Select(string.Format("{0} = '{1}'", columnaBusqueda, valorBusqueda));
                if (foundRows.Length > 0)
                {
                    return foundRows[0].ItemArray[columnaResultado].ToString();
                }
            }
            catch (Exception exc)
            {
            }
            return "";
        }

        public static int mtdGetConsecutivo(string Archivo)
        {
            int intConsecutivo = 1;
            StreamReader strSendReference = new StreamReader(Archivo);
            try
            {
                //Se crea variable para el send reference y se valida que el archivo no este vacio, si es asi se inicializa en 1

                string strSendreference = strSendReference.ReadToEnd();


                if (!strSendreference.IsNullOrWhiteSpace())
                {
                    intConsecutivo = Int16.Parse(strSendreference);
                }
                else
                {
                    intConsecutivo = 1;
                    
                }
                strSendReference.Close();
            }
            catch (FormatException exc)
            {
                intConsecutivo = 1;
                strSendReference.Close();
            }
            catch (Exception exc)
            {
                intConsecutivo = 1;
                strSendReference.Close();
                log.AddMessage(new GenericMessageLog { Message = string.Format("[Info] El archivo '{0}' no se pudo leer, se reinicia contador", Archivo) }, 0, string.Empty, exc.Message);
            }

            try
            {
                int intNuevoSendReference = intConsecutivo + 1;
                FileDirectory.WriteFile(intNuevoSendReference.ToString(), Archivo);
            }
            catch (Exception exc)
            {
                log.AddMessage(new GenericMessageLog { Message = string.Format("[Info] El archivo '{0}' no se pudo actualizar", Archivo) }, 0, string.Empty, exc.Message);
            }

            return intConsecutivo;
        }

        public static void mtdEnviarCorreo(string enviaCorreo, string servidor, string puerto, string usuario, string password, string autenticacion, string conexionSergura, string asunto, string quienEnvia, string destinatario, List<string> lstAdjuntos, LogEnvelope log)
        {

            try
            {
                if (enviaCorreo.Equals("1"))
                {
                    int? intPuerto = null;
                    if (!puerto.IsNullOrWhiteSpace())
                    {
                        intPuerto = int.Parse(puerto);
                    }

                    bool booAutenticacion = autenticacion.Equals("1");
                    bool booConexionSegura = conexionSergura.Equals("1");

                    SmtpWrapper mail = new SmtpWrapper
                    {
                        MailServer = servidor,
                        UserServer = @"" + usuario,
                        PasswordUserServer = @"" + password,
                        RequiresAuthentication = booAutenticacion,
                        EnableSSL = booConexionSegura,
                        Properties = new MailProperties
                        {
                            Subject = asunto,
                            From = quienEnvia,
                            To = destinatario,
                            Body = "Se adjuntan los archivos de LOGs generados.",
                            Attachments = lstAdjuntos
                        }
                    };
                    if (intPuerto.HasValue)
                    {
                        mail.Puerto = intPuerto.Value;
                    }
                    mail.Send();
                }
            }
            catch (FormatException e)
            {
                log.AddMessage(new GenericMessageLog { Message = string.Format("Falló al enviar E-Mail. Puerto '{0}' invalido", puerto) }, 0, string.Empty, e.Message);
            }
            catch (Exception e)
            {
                log.AddMessage(new GenericMessageLog { Message = "Falló al enviar E-Mail. Un parametro requerido es vacio" }, 0, string.Empty, e.Message);
            }
        }

        public static DataTable exceldata(string filePath, string nombreTabla, string hoja)
        {
            DataTable dtexcel = new DataTable(nombreTabla);
            bool hasHeaders = false;
            string HDR = hasHeaders ? "Yes" : "No";
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";

            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

            string query = "SELECT  * FROM [" + hoja + "$]";
            OleDbDataAdapter daexcel = new OleDbDataAdapter(query, conn);
            dtexcel.Locale = CultureInfo.CurrentCulture;
            daexcel.Fill(dtexcel);

            conn.Close();

            return dtexcel;

        }

        public static DateTime mtdGetFecha(string fecha)
        {
            DateTime dtFecha = new DateTime();
            try
            {
                int anio = int.Parse(fecha.Substring(0, 4));
                int mes = int.Parse(fecha.Substring(4, 2));
                int dia = int.Parse(fecha.Substring(6, 2));

                dtFecha = new DateTime(anio, mes, dia);
            }
            catch (Exception exc)
            {
            }
            return dtFecha;
        }

        public static bool mtdValidarFecha(string fecha)
        {
            try
            {
                int anio = int.Parse(fecha.Substring(0, 4));
                int mes = int.Parse(fecha.Substring(4, 2));
                int dia = int.Parse(fecha.Substring(6, 2));

                DateTime dtFecha = new DateTime(anio, mes, dia);
            }
            catch (Exception exc)
            {
                return false;
            }
            return true;
        }

        public static void mtdConcatenarArchivos(List<string> Archivos, string ruta)
        {
            StringBuilder sbContenidoArchivo = new StringBuilder();

            foreach (string archivo in Archivos)
            {
                sbContenidoArchivo.Append(File.ReadAllText(archivo));
            }

            File.WriteAllText(ruta, sbContenidoArchivo.ToString());
        }
    }
}