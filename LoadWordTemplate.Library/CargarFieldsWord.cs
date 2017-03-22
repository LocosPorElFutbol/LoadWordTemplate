using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace LoadWordTemplate.Library
{
    public class CargarFieldsWord
    {
        private string pathArchivo = string.Empty;

        public CargarFieldsWord(string _pathArchivo)
        {
            pathArchivo = _pathArchivo;
        }

        public void ReeplazarCampos()
        {
            try
            {
                if (string.IsNullOrEmpty(pathArchivo))
                    throw new Exception("Debe ingresar el nombre de un archivo");

                Object oMissing = System.Reflection.Missing.Value;
                Object oTemplatePath = pathArchivo;

                Application wordApp = new Application();
                Document wordDoc = new Document();

                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                foreach (Field myMergeField in wordDoc.Fields)
                {
                    Range rngFieldCode = myMergeField.Code;
                    String fieldText = rngFieldCode.Text;

                    // ONLY GETTING THE MAILMERGE FIELDS
                    if (fieldText.StartsWith(" MERGEFIELD"))
                    {
                        // THE TEXT COMES IN THE FORMAT OF
                        // MERGEFIELD  MyFieldName  \\* MERGEFORMAT
                        // THIS HAS TO BE EDITED TO GET ONLY THE FIELDNAME "MyFieldName"

                        Int32 endMerge = fieldText.IndexOf("\\");

                        Int32 fieldNameLength = fieldText.Length - endMerge;

                        String fieldName = fieldText.Substring(11, endMerge - 11);

                        // GIVES THE FIELDNAMES AS THE USER HAD ENTERED IN .dot FILE
                        fieldName = fieldName.Trim();

                        // **** FIELD REPLACEMENT IMPLEMENTATION GOES HERE ****//
                        // THE PROGRAMMER CAN HAVE HIS OWN IMPLEMENTATIONS HERE
                        if (fieldName == "NombreApellido14")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText("Salchi Jhon");
                        }
                        if (fieldName == "CP14")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText("1408");
                        }
                    }
                }
                wordDoc.SaveAs("myfile.doc");
                wordApp.Documents.Open("myFile.doc");
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }


    }
}
