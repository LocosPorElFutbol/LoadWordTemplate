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
    public class CumpleaniosCartaRepository
    {
        //Install-Package iTextSharp -Version 5.5.12

        private string pathSalidaDocumentoPdf = string.Empty;
        private const float ANCHO_TABLA = 0.0f;

        public CumpleaniosCartaRepository(string _pathSalidaDocumentoPdf)
        {
            this.pathSalidaDocumentoPdf = _pathSalidaDocumentoPdf;
        }

        public void GenerarCarta(IEnumerable<CartaEntity> ListaCartas, float wordSize)
        {
            try
            {
                // Creamos el documento con el tamaño de página tradicional
                Document doccumentoPdf = new Document(PageSize.LETTER);
                // Indicamos donde vamos a guardar el documento
                PdfWriter writer = PdfWriter.GetInstance(doccumentoPdf,
                                            new FileStream(this.pathSalidaDocumentoPdf, FileMode.Create));

                // Le colocamos el título y el autor
                // **Nota: Esto no será visible en el documento
                doccumentoPdf.AddCreator("LandManagement ©");

                // Abrimos el archivo
                doccumentoPdf.Open();

                foreach (CartaEntity carta in ListaCartas)
                {
                    PdfPTable tblPrueba = ArmarTablaCarta(carta, wordSize);
                    doccumentoPdf.Add(tblPrueba);
                    doccumentoPdf.Add(Chunk.NEXTPAGE);
                }

                doccumentoPdf.Close();
                writer.Close();

                System.Diagnostics.Process.Start(this.pathSalidaDocumentoPdf);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private PdfPTable ArmarTablaCarta(CartaEntity carta, float wordSize)
        {
            //iTextSharp.text.Font _standardFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            //iTextSharp.text.Font _standardFontEncabezado = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 13, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font _standardFontEncabezado = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, wordSize, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            PdfPCell cellColSpan = null;

            // Creamos una tabla que contendrá las columnas divisoras de la carta
            PdfPTable tablaCarta = new PdfPTable(3);
            tablaCarta.WidthPercentage = 100;
            //sacar bordes de la tabla principal
            tablaCarta.DefaultCell.Border = Rectangle.NO_BORDER;

            // Configuramos el título de las columnas de la tabla
            string fecha = carta.DiaCumpleanios + " " + carta.MesCumpleanios + " de " + DateTime.Now.Year;
			string caba = "C.A.B.A, ";
			PdfPCell cellDesde = new PdfPCell(new Phrase(string.Concat(caba, fecha), _standardFontEncabezado));
            cellDesde.Colspan = 3;
            cellDesde.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
            cellDesde.BorderWidth = ANCHO_TABLA;

            // Añadimos las celdas a la tabla
            tablaCarta.AddCell(cellDesde);

            // Configuramos el título de las columnas de la tabla
            //string fecha = carta.DiaCumpleanios + " " + carta.MesCumpleanios + " de " + DateTime.Now.Year;
            //PdfPCell cellFecha = new PdfPCell(new Phrase(fecha, _standardFontEncabezado));
            //cellFecha.Colspan = 3;
            //cellFecha.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
            //cellFecha.BorderWidth = ANCHO_TABLA;

            // Añadimos las celdas a la tabla
           // tablaCarta.AddCell(cellFecha);

            // Inserto linea en blanco
            PdfPCell cellLineaBlanco = null;
            cellLineaBlanco = new PdfPCell(new Phrase(Chunk.NEWLINE));
            cellLineaBlanco.Colspan = 3;
            cellLineaBlanco.Border = Rectangle.NO_BORDER;
            tablaCarta.AddCell(cellLineaBlanco);
            tablaCarta.AddCell(cellLineaBlanco);
            tablaCarta.AddCell(cellLineaBlanco);
            tablaCarta.AddCell(cellLineaBlanco);

            // Titulo y NombreApellido
            PdfPCell cellTitulo = new PdfPCell(new Phrase(carta.Titulo + " " + carta.NombreApellido, _standardFontEncabezado));
            cellTitulo.BorderWidth = ANCHO_TABLA;
            cellTitulo.Colspan = 2;
            tablaCarta.AddCell(cellTitulo);

            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.BorderWidth = ANCHO_TABLA;
            tablaCarta.AddCell(cellColSpan);

            // Direccion
            PdfPCell cellDireccion = new PdfPCell(new Phrase(carta.Direccion, _standardFontEncabezado));
            cellDireccion.BorderWidth = ANCHO_TABLA;
            cellDireccion.Colspan = 2;
            tablaCarta.AddCell(cellDireccion);

            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.BorderWidth = ANCHO_TABLA;
            tablaCarta.AddCell(cellColSpan);

            // CP, localidad y provincia
            PdfPCell cellLocalidadCP = new PdfPCell(new Phrase(carta.CPLocalidadProvincia, _standardFontEncabezado));
            cellLocalidadCP.BorderWidth = ANCHO_TABLA;
            cellLocalidadCP.Colspan = 2;
            tablaCarta.AddCell(cellLocalidadCP);

            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.BorderWidth = ANCHO_TABLA;
            tablaCarta.AddCell(cellColSpan);

            // Inserto 2 lineas en blanco
            tablaCarta.AddCell(cellLineaBlanco);
            tablaCarta.AddCell(cellLineaBlanco);

            // Saludo
            string saludo = "Hola " + carta.NombrePila + ":";
            PdfPCell cellHola = new PdfPCell(new Phrase(saludo, _standardFontEncabezado));
            cellHola.BorderWidth = ANCHO_TABLA;
            tablaCarta.AddCell(cellHola);

            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = ANCHO_TABLA;
            tablaCarta.AddCell(cellColSpan);

            // Cuerpo mensaje
            PdfPCell cellCuerpo = new PdfPCell(this.ObtenerParrafo(carta.CuerpoCarta, _standardFontEncabezado));
            cellCuerpo.Colspan = 3;

            cellCuerpo.BorderWidth = ANCHO_TABLA;
            tablaCarta.AddCell(cellCuerpo);

            // Inserto 3 lineas en blanco
            tablaCarta.AddCell(cellLineaBlanco);
            tablaCarta.AddCell(cellLineaBlanco);
            tablaCarta.AddCell(cellLineaBlanco);

            //dos celdas en blanco firma
            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = ANCHO_TABLA;
            tablaCarta.AddCell(cellColSpan);

            // Puntos firma
            PdfPCell cellPuntos = new PdfPCell(new Phrase(".............................................", _standardFontEncabezado));
            cellPuntos.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cellPuntos.BorderWidth = ANCHO_TABLA;
            tablaCarta.AddCell(cellPuntos);

            //dos celdas en blanco firma
            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = ANCHO_TABLA;
            tablaCarta.AddCell(cellColSpan);

            // Ruben
            PdfPCell cellRuben = new PdfPCell(new Phrase("RUBEN VIEYRA", _standardFontEncabezado));
            cellRuben.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cellRuben.BorderWidth = ANCHO_TABLA;
            tablaCarta.AddCell(cellRuben);

            //dos celdas en blanco firma
            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = ANCHO_TABLA;
            tablaCarta.AddCell(cellColSpan);

            // mensaje final
            //PdfPCell cellPie = new PdfPCell(new Phrase(this.GetPie(), _standardFontEncabezado));
            //cellPie.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            //cellPie.BorderWidth = ANCHO_TABLA;
            //tablaCarta.AddCell(cellPie);

            return tablaCarta;
        }

        private Paragraph ObtenerParrafo(string[] lineas, iTextSharp.text.Font fuente)
        {
            Paragraph parrafo = new Paragraph();
            parrafo.Font = fuente;
            try
            {
                bool primerlinea = true;

                foreach (string linea in lineas)
                {
                    if (string.IsNullOrEmpty(linea))
                    {
                        parrafo.Add(Chunk.NEWLINE);
                    }
                    else
                    {
                        string tab = linea.Substring(0, 1);
                        if (tab.Equals("\t"))
                        {
                            string lineaParrafo = linea.Substring(1);
                            if (primerlinea)
                            {
                                parrafo.Add(Chunk.CreateTabspace(75));
                                primerlinea = false;
                            }
                            else
                            {
                                parrafo.Add(Chunk.NEWLINE);
                                parrafo.Add(Chunk.CreateTabspace(75));
                            }
                            parrafo.Add(lineaParrafo);
                        }
                    }
                }

                return parrafo;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string GetPie()
        {
            return @"Pedile a Dios que renueve tus fuerzas y tu voluntad de cada día";
        }

    }
}
