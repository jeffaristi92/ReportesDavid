using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ProcesamientoXML
{
    public class Rutas
    {
        public string Archivo;
        public string Entrada;
        public string Salida;
        public string Backup;
        public string Referencias;
        public List<string> Adicionales;

        public Rutas()
        {
            Archivo = "";
            Entrada = "";
            Salida = "";
            Backup = "";
            Referencias = "";
            Adicionales = new List<string>();
        }

        public void cargarRutas()
        {
            string contenido = File.ReadAllText(Archivo);
            string[] lineas = contenido.Split('\n');
            if (lineas.Length >= 4)
            {
                Entrada = lineas[0].Replace('\r', ' ');
                Salida = lineas[1].Replace('\r', ' ');
                Backup = lineas[2].Replace('\r', ' ');
                Referencias = lineas[3].Replace('\r', ' ');

                if (lineas.Length > 4)
                {
                    for (int i = 4; i < lineas.Length; i++)
                    {
                        Adicionales.Add(lineas[i]);
                    }
                }
            }

            crearCarpetas();
        }

        public string crearRutas(string rutaBase)
        {
            StringBuilder contenidoArchivo = new StringBuilder();
            contenidoArchivo.AppendLine(Path.Combine(rutaBase, "Archivos_IN"));
            contenidoArchivo.AppendLine(Path.Combine(rutaBase, "Archivos_OUT"));
            contenidoArchivo.AppendLine(Path.Combine(rutaBase, "Backup"));
            contenidoArchivo.AppendLine(Path.Combine(rutaBase, "Referencias"));

            return contenidoArchivo.ToString();
        }

        public void actualizarRutas()
        {
            StringBuilder contenidoArchivo = new StringBuilder();
            contenidoArchivo.AppendLine(Entrada);
            contenidoArchivo.AppendLine(Salida);
            contenidoArchivo.AppendLine(Backup);
            contenidoArchivo.AppendLine(Referencias);

            File.Delete(Archivo);
            File.WriteAllText(contenidoArchivo.ToString(), Archivo);
        }

        bool crearCarpetas()
        {
            try
            {
                if (!Directory.Exists(Entrada))
                {
                    Directory.CreateDirectory(Entrada);
                }
                if (!Directory.Exists(Salida))
                {
                    Directory.CreateDirectory(Salida);
                }
                if (!Directory.Exists(Backup))
                {
                    Directory.CreateDirectory(Backup);
                }
                if (!Directory.Exists(Referencias))
                {
                    Directory.CreateDirectory(Referencias);
                }

                return true;
            }
            catch (Exception exc)
            {
                return false;
            }
        }
    }
}
