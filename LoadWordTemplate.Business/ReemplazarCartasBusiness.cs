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

        public void Reemplazar300Cartas()
        {
            try
            {
                reemplazarCartasRepository.Reemplazar300Cartas();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
