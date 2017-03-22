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
        private string pathTemplateWord = string.Empty;
        private string pathWordModificado = string.Empty;

        public ReemplazarEtiquetasRepository(string _pathTemplateWord, string _pathWordModificado)
        {
            pathTemplateWord = _pathTemplateWord;
            pathWordModificado = _pathWordModificado;
        }

        public void Reemplazar(List<CartaEntity> listaEtiquetas)
        {
            try
            {
                //OBJECT OF MISSING "NULL VALUE"
                Object oMissing = System.Reflection.Missing.Value;
                Object oTemplatePath = this.pathTemplateWord;

                Application wordApp = new Application();
                Document wordDoc = new Document();

                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
                //int i = 1;

                foreach (var obj in listaEtiquetas)
                {
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

                            switch (fieldName)
                            {
                                case "NombreCompletoApellido":
                                    myMergeField.Select();
                                    wordApp.Selection.TypeText(obj.NombreCompletoApellido);
                                    break;
                                case "Direccion":
                                    myMergeField.Select();
                                    wordApp.Selection.TypeText(obj.Direccion);
                                    break;
                                case "Localidad":
                                    myMergeField.Select();
                                    wordApp.Selection.TypeText(obj.Localidad);
                                    break;
                                case "CP":
                                    myMergeField.Select();
                                    wordApp.Selection.TypeText(obj.CodigoPostal);
                                    break;
                            }
                        }
                    }
                }

                wordDoc.SaveAs(pathWordModificado);
                wordDoc.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public void ReemplazarORIGINAL(List<CartaEntity> listaEtiquetas)
        {
            try
            {
                //OBJECT OF MISSING "NULL VALUE"
                Object oMissing = System.Reflection.Missing.Value;
                Object oTemplatePath = this.pathTemplateWord;

                Application wordApp = new Application();
                Document wordDoc = new Document();

                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                int i = 1;
                //Recorro la lista de etiquetas
                foreach (var obj in listaEtiquetas)
                {
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

                            // **** FIELD REPLACEMENT IMPLEMENTATION GOES HERE ****//
                            // THE PROGRAMMER CAN HAVE HIS OWN IMPLEMENTATIONS HERE
                            if (fieldName == "NombreApellido" + i.ToString())
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
                            if (fieldName == "CP" + i.ToString())
                            {
                                myMergeField.Select();
                                wordApp.Selection.TypeText(obj.CodigoPostal);
                            }
                        }
                    }
                    i++;
                }

                wordDoc.SaveAs(pathWordModificado);
                wordDoc.Close();
                //wordApp.Documents.Open("c:\\leo\\myFile.doc");
                //wordApp.Application.Quit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void ImprimirEtiquetas(double cantidadHojasAImprimir)
        {
            try
            {
                Application wordApp = new Application();
                wordApp.Visible = false;

                System.Windows.Forms.PrintDialog pDialog = new System.Windows.Forms.PrintDialog();
                if (pDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Document doc = wordApp.Documents.Add(pathWordModificado);

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

        /// <summary>
        /// Call the PrintOut method of the ThisDocument class in your project to print the entire document. 
        /// To use this example, run the code from the ThisDocument class.
        /// </summary>
        public void ImprimirEtiquetasMSDN()
        {
            //object copies = "1";
            //object pages = "";
            //object range = Word.WdPrintOutRange.wdPrintAllDocument;
            //object items = Word.WdPrintOutItem.wdPrintDocumentContent;
            //object pageType = Word.WdPrintOutPages.wdPrintAllPages;
            //object oTrue = true;
            //object oFalse = false;

            //this.PrintOut(ref oTrue, ref oFalse, ref range, ref missing, ref missing, ref missing,
            //    ref items, ref copies, ref pages, ref pageType, ref oFalse, ref oTrue,
            //    ref missing, ref oFalse, ref missing, ref missing, ref missing, ref missing);
        }
    }
}
