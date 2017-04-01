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

        public ReemplazarEtiquetasBusiness(string _pathTemplateWord, string _pathWordModificado)
        {
            reemplazarEtiquetasRepository = 
                new ReemplazarEtiquetasRepository(_pathTemplateWord, _pathWordModificado);
        }

        public void ReemplazarImprimir300Etiquetas(IEnumerable<CartaEntity> listaEtiquetas)
        {
            try
            {
                reemplazarEtiquetasRepository.ReemplazarImprimir300Etiquetas(listaEtiquetas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
