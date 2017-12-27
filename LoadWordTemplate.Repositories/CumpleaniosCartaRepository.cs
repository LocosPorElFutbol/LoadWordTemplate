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

        public CumpleaniosCartaRepository(string _pathSalidaDocumentoPdf)
        {
            this.pathSalidaDocumentoPdf = _pathSalidaDocumentoPdf;
        }

        public void GenerarCarta(IEnumerable<CartaEntity> ListaCartas)
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
                    PdfPTable tblPrueba = ArmarTablaCarta(carta);
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
        private PdfPTable ArmarTablaCarta(CartaEntity carta)
        {
            iTextSharp.text.Font _standardFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font _standardFontEncabezado = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            PdfPCell cellColSpan = null;

            // Creamos una tabla que contendrá las columnas divisoras de la carta
            PdfPTable tblPrueba = new PdfPTable(3);
            tblPrueba.WidthPercentage = 100;

            // Configuramos el título de las columnas de la tabla
            string fecha = carta.DiaCumpleanios + " de " + carta.MesCumpleanios + " de " + carta.AnioCumpleanios + ".";
            PdfPCell cellFecha = new PdfPCell(new Phrase("Ciudad Autónoma de Buenos Aires, " + fecha, _standardFontEncabezado));
            cellFecha.Colspan = 3;
            cellFecha.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
            cellFecha.BorderWidth = 0.1f;
            cellFecha.BorderWidthBottom = 0.1f;

            // Añadimos las celdas a la tabla
            tblPrueba.AddCell(cellFecha);

            // Inserto lineas en blanco
            for (int i = 0; i <= 2; i++)
            {
                PdfPCell cellLineaBlanco = null;
                cellLineaBlanco = new PdfPCell(new Phrase(Chunk.NEWLINE));
                cellFecha.Colspan = 3;
                cellFecha.BorderWidth = 0.1f;
                cellFecha.BorderWidthBottom = 0.1f;
                // Añadimos las celdas a la tabla
                tblPrueba.AddCell(cellLineaBlanco);
            }

            // Titulo
            PdfPCell cellTitulo = new PdfPCell(new Phrase(carta.Titulo, _standardFontEncabezado));
            cellTitulo.BorderWidth = 0.1f;
            cellTitulo.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellTitulo);

            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = 0.1f;
            cellColSpan.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellColSpan);

            // NombreApellido
            PdfPCell cellNombreApellido = new PdfPCell(new Phrase(carta.NombreCompletoApellido, _standardFontEncabezado));
            cellNombreApellido.BorderWidth = 0.1f;
            cellNombreApellido.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellNombreApellido);

            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = 0.1f;
            cellColSpan.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellColSpan);

            // Direccion
            PdfPCell cellDireccion = new PdfPCell(new Phrase(carta.Direccion, _standardFontEncabezado));
            cellDireccion.BorderWidth = 0.1f;
            cellDireccion.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellDireccion);

            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = 0.1f;
            cellColSpan.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellColSpan);

            // Localidad y CP
            PdfPCell cellLocalidadCP = new PdfPCell(new Phrase(carta.Localidad + " "  + carta.CodigoPostal, _standardFontEncabezado));
            cellLocalidadCP.BorderWidth = 0.1f;
            cellLocalidadCP.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellLocalidadCP);

            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = 0.1f;
            cellColSpan.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellColSpan);

            // Inserto 2 lineas en blanco
            for (int i = 0; i <= 1; i++)
            {
                PdfPCell cellLineaBlanco = null;
                cellLineaBlanco = new PdfPCell(new Phrase(Chunk.NEWLINE));
                cellLineaBlanco.Colspan = 3;
                cellFecha.BorderWidth = 0.1f;
                cellFecha.BorderWidthBottom = 0.1f;
                // Añadimos las celdas a la tabla
                tblPrueba.AddCell(cellLineaBlanco);
            }

            // Saludo
            string saludo = "Hola " + carta.NombrePila + ":";
            PdfPCell cellHola = new PdfPCell(new Phrase(saludo, _standardFontEncabezado));
            cellHola.BorderWidth = 0.1f;
            cellHola.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellHola);

            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = 0.1f;
            cellColSpan.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellColSpan);

            // Cuerpo mensaje
            //PdfPCell cellCuerpo = new PdfPCell(this.ObtenerParrafo("txbCuerpo.Text"));
            PdfPCell cellCuerpo = new PdfPCell(this.ObtenerParrafo(carta.CuerpoCarta));
            cellCuerpo.Colspan = 3;
            cellCuerpo.BorderWidth = 0.1f;
            cellCuerpo.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellCuerpo);

            // Inserto 3 lineas en blanco
            for (int i = 0; i <= 1; i++)
            {
                PdfPCell cellLineaBlanco = null;
                cellLineaBlanco = new PdfPCell(new Phrase(Chunk.NEWLINE));
                cellLineaBlanco.Colspan = 3;
                cellFecha.BorderWidth = 0.1f;
                cellFecha.BorderWidthBottom = 0.1f;
                // Añadimos las celdas a la tabla
                tblPrueba.AddCell(cellLineaBlanco);
            }

            //dos celdas en blanco firma
            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = 0.1f;
            cellColSpan.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellColSpan);

            // Puntos firma
            PdfPCell cellPuntos = new PdfPCell(new Phrase("....................................", _standardFontEncabezado));
            cellPuntos.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cellPuntos.BorderWidth = 0.1f;
            cellPuntos.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellPuntos);

            //dos celdas en blanco firma
            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = 0.1f;
            cellColSpan.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellColSpan);

            // Ruben
            PdfPCell cellRuben = new PdfPCell(new Phrase("RUBEN VIEYRA", _standardFontEncabezado));
            cellRuben.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cellRuben.BorderWidth = 0.1f;
            cellRuben.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellRuben);


            //dos celdas en blanco firma
            cellColSpan = new PdfPCell(new Phrase(string.Empty, _standardFontEncabezado));
            cellColSpan.Colspan = 2;
            cellColSpan.BorderWidth = 0.1f;
            cellColSpan.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellColSpan);

            // mensaje final
            PdfPCell cellPie = new PdfPCell(new Phrase(this.GetPie(), _standardFontEncabezado));
            cellPie.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cellPie.BorderWidth = 0.1f;
            cellPie.BorderWidthBottom = 0.1f;
            tblPrueba.AddCell(cellPie);
            return tblPrueba;
        }

        private Paragraph ObtenerParrafo(string[] lineas)
        {
            Paragraph parrafo = new Paragraph();
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
