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

        public ReemplazarCartasRepository(string _pathWordTemplateCarta)
        {
            pathWordTemplateCarta = _pathWordTemplateCarta;
        }

        public void AbrirTemplateCarta()
        {
            Object oMissing = System.Reflection.Missing.Value;
            Object oTemplatePath = this.pathWordTemplateCarta;

            Application wordApp = new Application();
            Document wordDoc = new Document();

            wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
            wordApp.Visible = true;
        }

        public void ReemplazarImprimir300Cartas(IEnumerable<CartaEntity> listaClientes)
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

                //Imprimir cartas del documento Actualizado
                this.ImprimirCartas(wordApp);

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
        /// Remplaza los campos del documento word por los valores que tiene la lista de clientes.
        /// </summary>
        /// <param name="wordApp">Aplicacion word que contiene el documento a modificar.</param>
        /// <param name="wordDoc">Documento Word que contiene los campos a actualizar</param>
        /// <param name="listaClientes">Lista de clientes que se utilizara para completar los campos del documento word.</param>
        /// <param name="cuerpoCarta">String que contiene el cuerpo de la carta.</param>
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
                            wordApp.Selection.TypeText(obj.NombreApellido);
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

        /// <summary>
        /// Imprime las cartas del documento word actualizado.
        /// </summary>
        /// <param name="wordApp">Application que contiene el documento word actualizado.</param>
        private void ImprimirCartas(Application wordApp)
        {
            try
            {
                System.Windows.Forms.PrintDialog pDialog = new System.Windows.Forms.PrintDialog();
                if (pDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Document worDoc = wordApp.Documents.Add(pathWordTemplateCarta300Actualizado);

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