using LoadWordTemplate.Entities;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;

namespace LoadWordTemplate.Repositories
{
    public class ReemplazarEtiquetasRepository
    {
        private string pathWordTemplateEtiquetas300 = string.Empty;
        private string pathWordTemplateEtiquetas300Actualizado = string.Empty;
        public int CANTIDAD_ETIQUETAS_POR_HOJA = 30;

        /// <summary>
        /// Constructor del repositorio responsable de reemplazar e imprimir las etiquetas de las cartas.
        /// </summary>
        /// <param name="_pathWordTemplateEtiquetas300">Path donde se encuentra el template del word a reeemplazar.</param>
        /// <param name="_pathWordTemplateEtiquetas300Actualizado">Path que contiene el archivo actualizado con la lista de clientes en las etiquetas.</param>
        public ReemplazarEtiquetasRepository
            (string _pathWordTemplateEtiquetas300, string _pathWordTemplateEtiquetas300Actualizado)
        {
            pathWordTemplateEtiquetas300 = _pathWordTemplateEtiquetas300;
            pathWordTemplateEtiquetas300Actualizado = _pathWordTemplateEtiquetas300Actualizado;
        }

        /// <summary>
        /// Metodo que reemplaza el template de etiquetas con los valores que contiene la lista que se envia como 
        /// parametro, posterior a esto ejecuta la orden de imprimir, dejando al usuario seleccionar la impresora
        /// deseada.
        /// </summary>
        /// <param name="listaClientes">Lista que contiene los datos de los clientes a completar en las etiquetas.</param>
        public void ReemplazarImprimir300Etiquetas(IEnumerable<CartaEntity> listaClientes)
        {
            try
            {
                //OBJECT OF MISSING "NULL VALUE"
                Object oMissing = System.Reflection.Missing.Value;
                Object oTemplatePath = this.pathWordTemplateEtiquetas300;

                Application wordApp = new Application();
                Document wordDoc = new Document();

                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                //Calculo la cantidad de hojas respecto de etiquetas
                int cantidadHojas = 
                    (int)Math.Ceiling((decimal)listaClientes.Count() / this.CANTIDAD_ETIQUETAS_POR_HOJA);

                //Eliminar secciones dependiendo de la cantidad de etiquetas
                this.EliminarSecciones(wordDoc, cantidadHojas);

                //Inserto etiquetas en blanco
                var lista = (List<CartaEntity>)listaClientes;
                int etiquetasEnBlanco = (this.CANTIDAD_ETIQUETAS_POR_HOJA * cantidadHojas) - listaClientes.Count();
                for (int j = 0; j < etiquetasEnBlanco; j++)
                    lista.Add(new CartaEntity());

                //Reemplazo valores de la lista en las etiquetas
                this.ReemplazarEtiquetas(wordApp, wordDoc, listaClientes);

                //Guardo como nuevo documento.
                wordDoc.SaveAs(pathWordTemplateEtiquetas300Actualizado);

                //Imprimo etiquetas
                this.ImprimirEtiquetas(wordApp);

                //Cierro documento word
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

        /// <summary>
        /// Reemplaza los campos de las etiquetas por los valores que contiene la lista de clientes.
        /// </summary>
        /// <param name="wordApp">Application word que contiene el documento actualizado.</param>
        /// <param name="wordDoc">Documento word actualizado.</param>
        /// <param name="listaClientes">Lista que contiene los datos de clientes a reflejar en las etiquetas.</param>
        private void ReemplazarEtiquetas(Application wordApp, Document wordDoc, IEnumerable<CartaEntity> listaClientes)
        {
            try
            {
                //Variable para recorrer etiquetas
                int i = 1;

                //Recorro la lista de etiquetas
                foreach (var obj in listaClientes)
                {
                    //Variable para detener el recorrido de campos cuando se completo la etiqueta
                    int j = 0;

                    //Recorro las etiquetas del WORD
                    foreach (Field myMergeField in wordDoc.Fields)
                    {
                        Range rngFieldCode = myMergeField.Code;
                        String fieldText = rngFieldCode.Text;

                        // ONLY GETTING THE MAILMERGE FIELDS
                        if (fieldText.StartsWith(" MERGEFIELD"))
                        {
                            //Metodo propio para obtener el nombre del filed (el de arriba pincha)
                            String fieldName = fieldText.Replace(" MERGEFIELD", "");

                            // GIVES THE FIELDNAMES AS THE USER HAD ENTERED IN .dot FILE
                            fieldName = fieldName.Trim();

                            if (fieldName == "NombreCompletoApellido" + i.ToString())
                            {
                                myMergeField.Select();
                                wordApp.Selection.TypeText(obj.NombreCompletoApellido);
                                j++;
                            }
                            if (fieldName == "Direccion" + i.ToString())
                            {
                                myMergeField.Select();
                                wordApp.Selection.TypeText(obj.Direccion);
                                j++;
                            }
                            if (fieldName == "Localidad" + i.ToString())
                            {
                                myMergeField.Select();
                                wordApp.Selection.TypeText(obj.Localidad);
                                j++;
                            }
                            if (fieldName == "CodigoPostal" + i.ToString())
                            {
                                myMergeField.Select();
                                wordApp.Selection.TypeText(obj.CodigoPostal);
                                j++;
                            }
                            if (j == 4)
                                break;
                        }
                    }
                    i++;
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Imprime las etiquetas del documento word actualizado.
        /// </summary>
        /// <param name="wordApp">Application que contiene el documento word actualizado.</param>
        private void ImprimirEtiquetas(Application wordApp)
        {
            try
            {
                System.Windows.Forms.PrintDialog pDialog = new System.Windows.Forms.PrintDialog();
                if (pDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Document worDoc = wordApp.Documents.Add(pathWordTemplateEtiquetas300Actualizado);

                    PrinterSettings printerSettings = pDialog.PrinterSettings;
                    wordApp.ActivePrinter = printerSettings.PrinterName;
                    wordApp.ActiveDocument.PrintOut();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
