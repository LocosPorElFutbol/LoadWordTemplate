using LoadWordTemplate.Business;
using LoadWordTemplate.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoadWordTemplate.Client
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var e = Etiquetas();
                //TestEtiquetas();
                TestCartas(e);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static List<CartaEntity> Etiquetas()
        {
            List<CartaEntity> lista = new List<CartaEntity>();

            CartaEntity e = new CartaEntity();
            e.Titulo = "Sr.";
            e.NombreCompleto = "Leonardo E. Choque Rodriguez";
            e.Direccion = "Calderon de la barca 2148, 7 E";
            e.Localidad = "CABA";
            e.CodigoPostal = "1407";
            e.FechaCumpleanios = new DateTime(1984, 4, 19);
            lista.Add(e);

            e = new CartaEntity();
            e.Titulo = "Sr.";
            e.NombreCompleto = "Don Jose de San Martin";
            e.Direccion = "alicante 1022";
            e.Localidad = "caba";
            e.CodigoPostal = "123";
            e.FechaCumpleanios = new DateTime(1778, 2, 25);
            lista.Add(e);

            e = new CartaEntity();
            e.Titulo = "Sra.";
            e.NombreCompleto = "bart simpson";
            e.Direccion = "siempre viva 742";
            e.Localidad = "springfield";
            e.CodigoPostal = "777";
            e.FechaCumpleanios = new DateTime(2000, 10, 25);
            lista.Add(e);

            //e = new CartaEntity();
            //e.NombreCompleto = "Homero Simpson";
            //e.Direccion = "Siempreviva 742";
            //e.Localidad = "Springfield";
            //e.CodigoPostal = "222";
            //lista.Add(e);

            //e = new CartaEntity();
            //e.NombreCompleto = "Pedro Picapiedra";
            //e.Direccion = "La edad de piedra";
            //e.Localidad = "NOSE";
            //e.CodigoPostal = "0600";
            //lista.Add(e);

            //e = new CartaEntity();
            //e.NombreCompleto = "Marulo Hernandez";
            //e.Direccion = "white 123";
            //e.Localidad = "CABA";
            //e.CodigoPostal = "1122";
            //lista.Add(e);

            return lista;
        }

        private static void TestEtiquetas()
        {
            var lista = Etiquetas();

            string pathTemplate = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\Template EtiquetasO.dotx";
            string pathNewWord = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\myfile.doc";

            ReemplazarEtiquetasBusiness reb = new ReemplazarEtiquetasBusiness(pathTemplate, pathNewWord);
            reb.Reemplazar(lista);
            double cantidadHojasImprimir = Math.Ceiling((double)lista.Count() / reb.CANTIDAD_ETIQUETAS_POR_HOJA);
            reb.ImprimirEtiquetas(cantidadHojasImprimir);

            //double cantidadHojasImprimir = Math.Ceiling((double)31 / reb.CANTIDAD_ETIQUETAS_POR_HOJA);
            Console.WriteLine(cantidadHojasImprimir.ToString());
            Console.ReadKey();
        }

        private static void TestCartas(List<CartaEntity> etiquetas)
        {
            try
            {
                string pathWordTemplateCarta = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\TemplateCarta.docx";
                string pathWordTemplateCarta300 = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\Secciones.docx";
//                string pathWordTemplateCarta300 = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\TemplateCarta300.docx";
                string pathWordTemplateCarta300Actualizado = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\TemplateCarta300Actualizado.docx";

                ReemplazarCartasBusiness rcb = new ReemplazarCartasBusiness(
                    pathWordTemplateCarta, pathWordTemplateCarta300, pathWordTemplateCarta300Actualizado);
                
                rcb.Reemplazar300Cartas(etiquetas);

                //double cantidadHojasImprimir = Math.Ceiling((double)31 / reb.CANTIDAD_ETIQUETAS_POR_HOJA);
                Console.WriteLine("Se ejecuto correctamente!!!");
                Console.ReadKey();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
