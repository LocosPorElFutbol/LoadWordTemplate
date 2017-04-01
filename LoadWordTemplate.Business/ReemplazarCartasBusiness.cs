using LoadWordTemplate.Entities;
using LoadWordTemplate.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoadWordTemplate.Business
{
    public class ReemplazarCartasBusiness
    {
        ReemplazarCartasRepository reemplazarCartasRepository = null;

        public ReemplazarCartasBusiness(
            string _pathWordTemplateCarta,
            string _pathWordTemplateCarta300,
            string _pathWordTemplateCarta300Actualizado)
        {
            reemplazarCartasRepository =
                new ReemplazarCartasRepository(
                    _pathWordTemplateCarta,
                    _pathWordTemplateCarta300,
                    _pathWordTemplateCarta300Actualizado);
        }

        public ReemplazarCartasBusiness(string _pathWordTemplateCarta)
        {
            reemplazarCartasRepository = new ReemplazarCartasRepository(_pathWordTemplateCarta);
        }

        public void AbrirTemplateCarta()
        {
            try
            {
                reemplazarCartasRepository.AbrirTemplateCarta();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void ReemplazarImprimir300Cartas(IEnumerable<CartaEntity> etiquetas)
        {
            try
            {
                reemplazarCartasRepository.ReemplazarImprimir300Cartas(etiquetas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
