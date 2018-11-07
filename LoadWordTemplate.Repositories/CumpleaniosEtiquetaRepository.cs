using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using LoadWordTemplate.Entities;

namespace LoadWordTemplate.Repositories
{
    public class CumpleaniosEtiquetaRepository
    {
        private const float ANCHO_TABLA = 0.0f;
        private const string SALTO_LINEA = "\r\n";

        private string pathDocumentoPdf = string.Empty;
        public CumpleaniosEtiquetaRepository(string _pathDocumentoPdf)
        {
            this.pathDocumentoPdf = _pathDocumentoPdf;
        }

        public void CrearEtiquetas(IEnumerable<CartaEntity> listaCartas)
        {
            try
            {
                listaCartas = ValidarCantidadEtiquetas(listaCartas);

                // Creamos el documento con el tamaño de página tradicional
                Document doccumentoPdf = new Document(PageSize.LETTER, 14.17f, 14.17f, 34.02f, 34.02f);
                // Indicamos donde vamos a guardar el documento
                PdfWriter writer = PdfWriter.GetInstance(doccumentoPdf, new FileStream(this.pathDocumentoPdf, FileMode.Create));

                // Le colocamos el título y el autor
                // **Nota: Esto no será visible en el documento
                doccumentoPdf.AddCreator("LandManagement ©");

                // Abrimos el archivo
                doccumentoPdf.Open();

                //Defino fuentes
                iTextSharp.text.Font _standardFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);

                // Creamos una tabla que contendrá las etiquetas en 3 columnas 
                PdfPTable tablaEtiquetas = new PdfPTable(3);
                tablaEtiquetas.WidthPercentage = 100;

                //************************* Tabla de etiquetas *************************
                PdfPCell cellEtiqueta = null;

                foreach (CartaEntity carta in listaCartas)
                {
                    string textoEtiqueta = carta.NombreApellido + SALTO_LINEA;
                    textoEtiqueta += carta.Direccion + SALTO_LINEA;
                    textoEtiqueta += carta.CPLocalidadProvincia;

                    //Agrego filas en la tabla de etiquetas
                    cellEtiqueta = new PdfPCell(new Phrase(textoEtiqueta, _standardFont));
                    cellEtiqueta.Padding = 8.5f;
                    cellEtiqueta.FixedHeight = 72.01f;//25.49mm
                    cellEtiqueta.BorderWidth = ANCHO_TABLA;

                    tablaEtiquetas.AddCell(cellEtiqueta);
                }

                // Finalmente, añadimos la tabla al documento PDF y cerramos el documento
                doccumentoPdf.Add(tablaEtiquetas);

                doccumentoPdf.Close();
                writer.Close();

                System.Diagnostics.Process.Start(this.pathDocumentoPdf);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        private IEnumerable<CartaEntity> ValidarCantidadEtiquetas(IEnumerable<CartaEntity> listaCartas)
        {
            List<CartaEntity> lista = (List<CartaEntity>)listaCartas;

            double promedio = (double)listaCartas.Count() / 3;
            int etiquetasMultiploDeTres = (int)Math.Ceiling(promedio) * 3;
            int cantidadDeEtiquetasAAgregar = etiquetasMultiploDeTres - listaCartas.Count();

            for (int i = 0; i < cantidadDeEtiquetasAAgregar; i++)
                lista.Add(new CartaEntity());

            return lista;
        }

        /// <summary>
        /// Convierte a valores float la medida en milimetros (no se esta utilizando, en caso de requerirlo, aplicarlo.)
        /// </summary>
        /// <param name="milimetros">Valor en milimetros a convertir</param>
        /// <returns>float equivalente a la medida pdfpcell</returns>
        private float ConvertMMtoFloatPdfpSize(int milimetros)
        {
            // Referencia de medidas f-mm: 70.0f = 24.69mm
            float mmResult = (milimetros * 70) / 24.69f;
            return float.Parse(mmResult.ToString("F"));
        }
    }
}
