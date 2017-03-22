using LoadWordTemplate.Entities;
using LoadWordTemplate.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoadWordTemplate.Business
{
    public class ReemplazarEtiquetasBusiness
    {
        private ReemplazarEtiquetasRepository reemplazarEtiquetasRepository = null;
        public int CANTIDAD_ETIQUETAS_POR_HOJA = 30;

        public ReemplazarEtiquetasBusiness(string _pathTemplateWord, string _pathWordModificado)
        {
            reemplazarEtiquetasRepository = 
                new ReemplazarEtiquetasRepository(_pathTemplateWord, _pathWordModificado);
        }

        public void Reemplazar(List<CartaEntity> listaEtiquetas)
        {
            try
            {
                reemplazarEtiquetasRepository.Reemplazar(listaEtiquetas);
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
                reemplazarEtiquetasRepository.ImprimirEtiquetas(cantidadHojasAImprimir);
            }
            catch (Exception ex)
            {
                throw ex;
            }        
        }
    }
}
