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
                Document doccumentoPdf = new Document(PageSize.LETTER, 0f, 0f, 0f, 0f);
                // Indicamos donde vamos a guardar el documento
                PdfWriter writer = PdfWriter.GetInstance(doccumentoPdf,
                                            new FileStream(this.pathDocumentoPdf, FileMode.Create));

                // Le colocamos el título y el autor
                // **Nota: Esto no será visible en el documento
                doccumentoPdf.AddCreator("LandManagement ©");

                // Abrimos el archivo
                doccumentoPdf.Open();

                //Defino fuentes
                iTextSharp.text.Font _standardFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                iTextSharp.text.Font _standardFontEncabezado = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 1, iTextSharp.text.Font.NORMAL, BaseColor.WHITE);

                // Creamos una tabla que contendrá las etiquetas en 3 columnas 
                PdfPTable tablaEtiquetas = new PdfPTable(3);
                tablaEtiquetas.WidthPercentage = 100;
                //Sacar bordes de la tabla principal
                tablaEtiquetas.DefaultCell.Border = Rectangle.NO_BORDER;

                // Configuramos el título de las columnas de la tabla
                PdfPCell columna1 = new PdfPCell(new Phrase("", _standardFontEncabezado));
                columna1.BorderWidth = ANCHO_TABLA;
                //columna1.BorderWidthBottom = ANCHO_TABLA;

                PdfPCell columna2 = new PdfPCell(new Phrase("", _standardFontEncabezado));
                columna2.BorderWidth = ANCHO_TABLA;
                //columna2.BorderWidthBottom = ANCHO_TABLA;

                PdfPCell columna3 = new PdfPCell(new Phrase("", _standardFontEncabezado));
                columna3.BorderWidth = ANCHO_TABLA;
                //columna3.BorderWidthBottom = ANCHO_TABLA;

                // Añadimos las celdas a la tabla
                tablaEtiquetas.AddCell(columna1);
                tablaEtiquetas.AddCell(columna2);
                tablaEtiquetas.AddCell(columna3);

                //************************* Tabla de etiquetas *************************
                PdfPTable tablaEtiqueta = null;
                PdfPCell cellNombre = null;
                PdfPCell cellApellido = null;
                PdfPCell cellDireccion = null;
                PdfPCell cellCodigoPostal = null;

                foreach (CartaEntity carta in listaCartas)
                {
                    tablaEtiqueta = new PdfPTable(1);
                    tablaEtiqueta.WidthPercentage = 100;

                    //Agrego filas en la tabla de etiquetas
                    cellNombre = new PdfPCell(new Phrase(carta.Nombre, _standardFont));
                    cellNombre.BorderWidth = ANCHO_TABLA;

                    cellApellido = new PdfPCell(new Phrase(carta.Apellido, _standardFont));
                    cellApellido.BorderWidth = ANCHO_TABLA;

                    cellDireccion = new PdfPCell(new Phrase(carta.Direccion, _standardFont));
                    cellDireccion.BorderWidth = ANCHO_TABLA;

                    cellCodigoPostal = new PdfPCell(new Phrase(carta.CodigoPostal + carta.Localidad + carta.Provincia, _standardFont));
                    cellCodigoPostal.BorderWidth = ANCHO_TABLA;

                    tablaEtiqueta.AddCell(cellNombre);
                    tablaEtiqueta.AddCell(cellApellido);
                    tablaEtiqueta.AddCell(cellDireccion);
                    tablaEtiqueta.AddCell(cellCodigoPostal);

                    tablaEtiquetas.AddCell(tablaEtiqueta);
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

    }
}
