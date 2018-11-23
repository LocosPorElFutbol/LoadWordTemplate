using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace LoadWordTemplate.Entities
{
    public class CartaEntity
    {
        private DateTime _FechaCumpleanios;
        public DateTime FechaCumpleanios
        { get { return _FechaCumpleanios; } set { _FechaCumpleanios = value; } }

        private string _Titulo = string.Empty;
        public string Titulo { get { return _Titulo; } set { _Titulo = value; } }

        private string _Nombre = string.Empty;
        public string Nombre { get { return _Nombre; } set { _Nombre = value; } }

        private string _Apellido = string.Empty;
        public string Apellido { get { return _Apellido; } set { _Apellido = value; } }

        private string _Direccion = string.Empty;
        public string Direccion { get { return _Direccion; } set { _Direccion = value; } }

        private string _Localidad = string.Empty;
        public string Localidad { get { return _Localidad; } set { _Localidad = value; } }

        private string _Provincia = string.Empty;
        public string Provincia { get { return _Provincia; } set { _Provincia = value; } }

        private string _CodigoPostal = string.Empty;
        public string CodigoPostal { get { return _CodigoPostal; } set { _CodigoPostal = value; } }

        private string _NombrePila = string.Empty;
        public string NombrePila { get { return _NombrePila; } set { _NombrePila = value; } }

        private string[] _CuerpoCarta = null;
        public string[] CuerpoCarta { get { return _CuerpoCarta; } set { _CuerpoCarta = value; } }

        public string NombreApellido
        {
            get { return Nombre + " " + Apellido; }
        }

        public string CPLocalidadProvincia
        {
            get
            {
                if (string.IsNullOrEmpty(CodigoPostal) && string.IsNullOrEmpty(Localidad) && string.IsNullOrEmpty(Provincia))
                    return string.Empty;

                if (Localidad.ToUpper().Equals("CABA"))
                    return string.Format("CP {0}, {1}", CodigoPostal, Localidad);
                else
                    return string.Format("CP {0}, {1}, {2}", CodigoPostal, Localidad, Provincia);
            }
        }

        public string DiaCumpleanios
        {
            get { return FechaCumpleanios.Day.ToString(); }
        }
        public string MesCumpleanios
        {
            get { return FechaCumpleanios.ToString("MMMM", CultureInfo.CreateSpecificCulture("es")); }
        }
        public string AnioCumpleanios
        {
            get { return FechaCumpleanios.Year.ToString(); }
        }

        /// <summary>
        /// Este constructor se utiliza para las etiquetas que no son multiplo de 3.
        /// </summary>
        public CartaEntity()
        {
            _Titulo = string.Empty;
            _NombrePila = string.Empty;
            _Nombre = string.Empty;
            _Apellido = string.Empty;
            _Direccion = string.Empty;
            _Localidad = string.Empty;
            _Provincia = string.Empty;
            _CodigoPostal = string.Empty;
            _FechaCumpleanios = new DateTime(1900, 01, 01);
            _CuerpoCarta = new string[0];
        }

        public CartaEntity(string pTitulo,
            string pNombrePila,
            string pNombre,
            string pApellido,
            string pDireccion,
            string pLocalidad,
            string pProvincia,
            string pCodigoPostal,
            DateTime pFechaCumpleanios,
            string[] pCuerpoCarta)
        {
            _Titulo = pTitulo;
            _NombrePila = pNombrePila;
            _Nombre = pNombre;
            _Apellido = pApellido;
            _Direccion = pDireccion;
            _Localidad = pLocalidad;
            _Provincia = pProvincia;
            _CodigoPostal = pCodigoPostal;
            _FechaCumpleanios = pFechaCumpleanios;
            _CuerpoCarta = pCuerpoCarta;
        }
    }
}

