using LoadWordTemplate.Entities;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;

namespace LoadWordTemplate.Repositories
{
    public class ReemplazarCartasRepository
    {
        private string pathWordTemplateCarta = string.Empty;
        private string pathWordTemplateCarta300 = string.Empty;
        private string pathWordTemplateCarta300Actualizado = string.Empty;
        private string textoCuerpo = string.Empty;

        public ReemplazarCartasRepository(
            string _pathWordTemplateCarta, 
            string _pathWordTemplateCarta300,
            string _pathWordTemplateCarta300Actualizado)
        {
            pathWordTemplateCarta = _pathWordTemplateCarta;
            pathWordTemplateCarta300 = _pathWordTemplateCarta300;
            pathWordTemplateCarta300Actualizado = _pathWordTemplateCarta300Actualizado;
        }

        public void Reemplazar300Cartas(IEnumerable<CartaEntity> listaClientes)
        {
            try
            {
                //Obtengo el cuerpo del mensaje
                string cuerpoCarta = this.ObtenerCuerpoCarta();

                //OBJECT OF MISSING "NULL VALUE"
                Object oMissing = System.Reflection.Missing.Value;
                Object oTemplatePath = this.pathWordTemplateCarta300;

                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                Document wordDoc = new Document();

                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                //Elimino las secciones del documento que no voy a utilizar
                this.EliminarSecciones(wordDoc, listaClientes.Count());

                //Reemplazo los campos por los valores de la lista
                this.ReemplazarCampos(wordApp, wordDoc, listaClientes, cuerpoCarta);

                //Guardo los cambios en nuevo archivo word
                wordDoc.SaveAs(pathWordTemplateCarta300Actualizado);
                ((Microsoft.Office.Interop.Word._Document)wordDoc).Close();

                //Cierro la aplicación word
                object oMissingValue = System.Reflection.Missing.Value;
                ((_Application)wordApp).Quit(oMissingValue, oMissingValue, oMissingValue);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Elimina las secciones de un documento word a partir de la siguiente pagina que se envia como parámetro.
        /// </summary>
        /// <param name="doc">Documento word a modificar</param>
        /// <param name="borrarPagina">ultima pagina que conservará el documento, se eliminarán a partir de la siguiente a este parametro.</param>
        private void EliminarSecciones(Document doc, int borrarPagina)
        {
            object missing = Type.Missing;

            foreach (Microsoft.Office.Interop.Word.Section section in doc.Sections)
            {
                if (section.Index > borrarPagina)
                    section.Range.Delete(ref missing, ref missing);
            }
        }

        private void ReemplazarCampos(Application wordApp, Document wordDoc, IEnumerable<CartaEntity> listaClientes, string cuerpoCarta)
        {
            int i = 1;
            //Recorro la lista de etiquetas
            foreach (var obj in listaClientes)
            {
                //Recorro las etiquetas del WORD
                foreach (Field myMergeField in wordDoc.Fields)
                {
                    Range rngFieldCode = myMergeField.Code;
                    String fieldText = rngFieldCode.Text;

                    // ONLY GETTING THE MAILMERGE FIELDS
                    if (fieldText.StartsWith(" MERGEFIELD"))
                    {
                        //Extraigo el mergedfield del word, lo spliteo porque word
                        //ingresa caracteres raros al modificar el texto del mergedfield
                        String fieldNameCompuesto = fieldText.Replace(" MERGEFIELD", "");
                        string[] arrayCampo = fieldNameCompuesto.Split('\\');
                        string fieldName = arrayCampo[0].ToString();

                        //Elimino espacios en blanco
                        fieldName = fieldName.Trim();

                        //Reemplazo mi mergedmail con el texto del WordTemplateCarta
                        if (fieldName == "DiaCumpleanios" + i.ToString())
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(obj.DiaCumpleanios);
                        }
                        if (fieldName == "Mes" + i.ToString())
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(obj.MesCumpleanios);
                        }
                        if (fieldName == "Anio" + i.ToString())
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(DateTime.Now.Year.ToString());
                        }
                        if (fieldName == "Titulo" + i.ToString())
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(obj.Titulo);
                        }
                        if (fieldName == "NombreCompletoApellido" + i.ToString())
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(obj.NombreCompletoApellido);
                        }
                        if (fieldName == "Direccion" + i.ToString())
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(obj.Direccion);
                        }
                        if (fieldName == "Localidad" + i.ToString())
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(obj.Localidad);
                        }
                        if (fieldName == "CodigoPostal" + i.ToString())
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(obj.CodigoPostal);
                        }
                        if (fieldName == "NombrePila" + i.ToString())
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(obj.NombrePila);
                        }
                        if (fieldName == "CuerpoCarta" + i.ToString())
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(cuerpoCarta);
                        }
                    }
                }
                i++;
            }
        }

        /// <summary>
        /// Obtiene el cuerpo de la carta, que se encuentra en los tags "« »". Nombre del campo: CuerpoCarta.
        /// </summary>
        /// <returns>Retorna un string que contiene el texto de la carta.</returns>
        private string ObtenerCuerpoCarta()
        {
            try
            {
                //OBJECT OF MISSING "NULL VALUE"
                Object oMissing = System.Reflection.Missing.Value;
                Object oTemplatePath = this.pathWordTemplateCarta;

                Application wordApp = new Application();
                Document wordDoc = new Document();

                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                //Recorro las etiquetas de template carta
                foreach (Field myMergeField in wordDoc.Fields)
                {
                    Range rngFieldCode = myMergeField.Code;
                    String fieldText = rngFieldCode.Text;

                    // obtengo solo los mergedfield
                    if (fieldText.StartsWith(" MERGEFIELD"))
                    {
                        //Metodo propio para obtener el nombre del mergedfiled
                        String fieldNameCompuesto = fieldText.Replace(" MERGEFIELD", "");
                        string[] arrayCampo = fieldNameCompuesto.Split('\\');
                        string fieldName = arrayCampo[0].ToString();

                        // Elimino espacios en blanco
                        fieldName = fieldName.Trim();

                        //Obtengo el cuerpo de la carta y elimino los tags « »
                        if (fieldName == "CuerpoCarta")
                        {
                            myMergeField.Select();
                            textoCuerpo = wordApp.Selection.Text;
                            textoCuerpo = textoCuerpo.Replace("«", "");
                            textoCuerpo = textoCuerpo.Replace("»", "");
                            break;
                        }
                    }
                }

                object oMissingValue = System.Reflection.Missing.Value;
                ((_Application)wordApp).Quit(oMissingValue, oMissingValue, oMissingValue);

                return textoCuerpo;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void ImprimirCartas(double cantidadHojasAImprimir)
        {
            try
            {
                Application wordApp = new Application();
                wordApp.Visible = false;

                System.Windows.Forms.PrintDialog pDialog = new System.Windows.Forms.PrintDialog();
                if (pDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Document doc = wordApp.Documents.Add(pathWordTemplateCarta300);

                    PrinterSettings printerSettings = pDialog.PrinterSettings;
                    wordApp.ActivePrinter = printerSettings.PrinterName;

                    //wordApp.ActivePrinter = pDialog.PrinterSettings.PrinterName;                

                    wordApp.ActiveDocument.PrintOut(); //this will also work: doc.PrintOut();
                    doc.Close(SaveChanges: false);
                    doc = null;

                    object copies = "1";
                    object pages = "1";
                    object range = Microsoft.Office.Interop.Word.WdPrintOutRange.wdPrintCurrentPage;
                    object items = Microsoft.Office.Interop.Word.WdPrintOutItem.wdPrintDocumentContent;
                    object pageType = Microsoft.Office.Interop.Word.WdPrintOutPages.wdPrintAllPages;
                    object oTrue = true;
                    object oFalse = false;
                    object missing = System.Reflection.Missing.Value;
                    //Word.Document document = this.Application.ActiveDocument;
                    Microsoft.Office.Interop.Word.Document document = doc;

                    document.PrintOut(
                        ref oTrue, ref oFalse, ref range, ref missing, ref missing, ref missing,
                        ref items, ref copies, ref pages, ref pageType, ref oFalse, ref oTrue,
                        ref missing, ref oFalse, ref missing, ref missing, ref missing, ref missing);


                    //IMPRIMIR
                    //https://msdn.microsoft.com/en-us/library/b9f0ke7y.aspx
                    //http://stackoverflow.com/questions/878302/printing-using-word-interop-with-print-dialog
                    //object copies = "1";
                    //object pages = "1";
                    //object range = Word.WdPrintOutRange.wdPrintCurrentPage;
                    //object items = Word.WdPrintOutItem.wdPrintDocumentContent;
                    //object pageType = Word.WdPrintOutPages.wdPrintAllPages;
                    //object oTrue = true;
                    //object oFalse = false;
                    //Word.Document document = this.Application.ActiveDocument;

                    //document.PrintOut(
                    //    ref oTrue, ref oFalse, ref range, ref missing, ref missing, ref missing,
                    //    ref items, ref copies, ref pages, ref pageType, ref oFalse, ref oTrue,
                    //    ref missing, ref oFalse, ref missing, ref missing, ref missing, ref missing);
                }

                // <EDIT to include Jason's suggestion>
                ((_Application)wordApp).Quit(SaveChanges: false);
                // </EDIT>

                // Original: wordApp.Quit(SaveChanges: false);
                wordApp = null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}