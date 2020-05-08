using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LoadWordTemplate.Entities;
using LoadWordTemplate.Repositories;

namespace LoadWordTemplate.Business
{
    public class CumpleaniosEtiquetaBusiness
    {
        private CumpleaniosEtiquetaRepository cumpleaniosEtiquetaRepository = null;
        public CumpleaniosEtiquetaBusiness(string _pathEtiquetasPdf)
        {
            this.cumpleaniosEtiquetaRepository = new CumpleaniosEtiquetaRepository(_pathEtiquetasPdf);
        }

        public void CrearEtiquetas(IEnumerable<CartaEntity> listaCartas, float wordSize)
        {
            this.cumpleaniosEtiquetaRepository.CrearEtiquetas(listaCartas, wordSize);
        }
    }
}
