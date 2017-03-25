using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace LoadWordTemplate.Entities
{
    public class CartaEntity
    {
        public DateTime FechaCumpleanios { get; set; }
        public string Titulo { get; set; }
        public string NombreCompleto { get; set; }
        public string Apellido { get; set; }
        public string Direccion { get; set; }
        public string Localidad { get; set; }
        public string CodigoPostal { get; set; }
        public string NombrePila { get; set; }
        public string CuerpoCarta { get; set; }
        public string NombreCompletoApellido
        {
            get { return NombreCompleto + " " + Apellido; }
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
    }
}

