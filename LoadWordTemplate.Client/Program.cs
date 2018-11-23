using LoadWordTemplate.Business;
using LoadWordTemplate.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoadWordTemplate.Client
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var e = Etiquetas();
                //OLD
                //TestEtiquetas();
                //TestAbrirTemplateCarta();
                //TestReemplazar300Cartas(e);

                //NEW
                //CrearCartas(e);
                CrearEtiquetas(e);
                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static List<CartaEntity> Etiquetas()
        {
            string cuerpoCarta = "\tEs bueno observar la actitud de los pájaros ante la adversidad, pasan días y días haciendo su nido y recogiendo materiales, muchos de estos traídos desde largas distancias, y cuando ya está terminado y listo para poner los huevos, las inclemencias del tiempo, la mano del hombre o la obra de algún animal, destruye y tira por el suelo lo que con tanto esfuerzo se logró.  ¿Qué hace el pájaro?  ¿Se lamenta, se paraliza, abandona la tarea?  … De ninguna manera!!!  VUELVE A EMPEZAR, una y otra vez hasta que en el nido aparecen los primeros huevos.  A veces, muchas veces, antes de que nazcan los pichones, algún animal o una tormenta vuelve a destruir el nido, pero esta vez con su precioso contenido. Aún así el pájaro jamás retrocede, sigue construyendo, y nunca deja de cantar.\r\n\r\n\tHoy empieza un nuevo año en tu vida ¿Sentiste alguna vez que tu vida, tu trabajo, tu familia, tus amigos no son lo que soñaste?  ¿Te dieron ganas de decir ¡Basta!, no vale la pena el esfuerzo, esto es demasiado para mí?  ¿Muchas veces te cansaste de volver a empezar, del desgaste de la lucha diaria, de la confianza traicionada, de las metas no alcanzadas cuando estabas a punto de lograrlo?\r\n\r\n\tPor más que la vida te golpee, no te entregues nunca.  No te preocupes si en la batalla sufrís alguna herida, es de esperar que algo así suceda.  Junta los pedazos de tu esperanza, ármala de nuevo y volvé a empezar.  No importa lo que pase, no aflojes, dale para adelante.  La vida es un desafío constante, pero vale la pena aceptarlo y sobre todo NUNCA DEJES DE CANTAR.";

            List<CartaEntity> lista = new List<CartaEntity>();

            CartaEntity e = new CartaEntity();
            e.Titulo = "Sr.";
            e.NombrePila = "Leito";
            e.Nombre = "Leonardo Elvio";
            e.Apellido = "Choque Rodriguez";
            e.Provincia = "Santiago del Estero.";
            e.Direccion = "Calderon de la barca 2148, 7 E";
            e.Localidad = "Alguno que sea recontra requete largo para ver que onda.";
            e.CodigoPostal = "1407";
            e.FechaCumpleanios = new DateTime(1984, 4, 19);
            e.CuerpoCarta = new[] { cuerpoCarta };
            lista.Add(e);

            e = new CartaEntity();
            e.Titulo = "Sr.";
            e.NombrePila = "José";
            e.Nombre = "Don Jose";
            e.Apellido = "de San Martin";
            e.Direccion = "alicante 1022";
            e.Localidad = "caba";
            e.Provincia = "Ciudad autónoma de buenos aires.";
            e.CodigoPostal = "123";
            e.FechaCumpleanios = new DateTime(1778, 2, 25);
            e.CuerpoCarta = new[] { cuerpoCarta };
            lista.Add(e);

            e = new CartaEntity();
            e.Titulo = "Sra.";
            e.NombrePila = "Barto";
            e.Nombre = "bartolomeo";
            e.Apellido = "J Simpson";
            e.Direccion = "siempre viva 742";
            e.Provincia = "Ciudad autónoma de buenos aires.";
            e.Localidad = "springfield";
            e.CodigoPostal = "777";
            e.FechaCumpleanios = new DateTime(2000, 10, 25);
            e.CuerpoCarta = new[] { cuerpoCarta };
            lista.Add(e);

            e = new CartaEntity();
            e.Titulo = "Sra.";
            e.Nombre = "Homero J";
            e.Apellido = "Simpson";
            e.Direccion = "Siempreviva 742";
            e.Provincia = "Ciudad autónoma de buenos aires.";
            e.Localidad = "Springfield";
            e.CodigoPostal = "222";
            e.FechaCumpleanios = new DateTime(2000, 10, 25);
            e.CuerpoCarta = new[] { cuerpoCarta };
            lista.Add(e);

            e = new CartaEntity();
            e.Titulo = "Sra.";
            e.Nombre = "Pedro";
            e.Apellido = "Picapiedra";
            e.Direccion = "La edad de piedra";
            e.Localidad = "NOSE";
            e.Provincia = "Cordoba";
            e.CodigoPostal = "0600";
            e.FechaCumpleanios = new DateTime(2000, 10, 15);
            e.CuerpoCarta = new[] { cuerpoCarta };
            lista.Add(e);

            e = new CartaEntity();
            e.Titulo = "Sr.";
            e.Nombre = "Marulo";
            e.Apellido = "Hernandez";
            e.Direccion = "white 123";
            e.Localidad = "CABA";
            e.Provincia = "Ciudad autónoma de buenos aires.";
            e.CodigoPostal = "1122";
            e.FechaCumpleanios = new DateTime(2013, 07, 22);
            e.CuerpoCarta = new[] { cuerpoCarta };
            lista.Add(e);

            for (int i = 0; i < 50; i++)
            {
                e = new CartaEntity();
                e.Titulo = "Sr.";
                e.Nombre = "Jonhy";
                e.Apellido = "Melaslavo";
                e.Direccion = "calle falsa 321";
                e.Localidad = "CABA";
                e.CodigoPostal = "1415";
                e.FechaCumpleanios = new DateTime(2013, 07, 22);
                e.CuerpoCarta = new[] { cuerpoCarta };
                lista.Add(e);
            }
            return lista;
        }

        private static void TestEtiquetas()
        {
            var lista = Etiquetas();

            string pathTemplate = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\TemplateEtiquetas300.docx";
            string pathTemplateActualizado = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\TemplateEtiquetas300Actualizado.docx";

            ReemplazarEtiquetasBusiness reb = new ReemplazarEtiquetasBusiness(pathTemplate, pathTemplateActualizado);
            reb.ReemplazarImprimir300Etiquetas(lista);

            Console.WriteLine("Se ejecuto correctamente!!!");
            Console.ReadKey();
        }

        private static void TestReemplazar300Cartas(List<CartaEntity> etiquetas)
        {
            try
            {
                string pathWordTemplateCarta = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\TemplateCarta.docx";
                string pathWordTemplateCarta300 = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\TemplateCarta300.docx";
                string pathWordTemplateCarta300Actualizado = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\TemplateCarta300Actualizado.docx";

                ReemplazarCartasBusiness rcb = new ReemplazarCartasBusiness(
                    pathWordTemplateCarta, pathWordTemplateCarta300, pathWordTemplateCarta300Actualizado);

                rcb.ReemplazarImprimir300Cartas(etiquetas);

                Console.WriteLine("Se ejecuto correctamente!!!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void TestAbrirTemplateCarta()
        {
            string pathWordTemplateCarta = "C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\TemplateWord\\Documentos\\Pruebas\\TemplateCarta.docx";

            ReemplazarCartasBusiness cartas = new ReemplazarCartasBusiness(pathWordTemplateCarta);
            cartas.AbrirTemplateCarta();
        }

        private static void CrearCartas(IEnumerable<CartaEntity> cartas)
        {
            try
            {
                CumpleaniosCartaBusiness cartaCumpleaniosBusiness = new CumpleaniosCartaBusiness("C:\\Leo\\pruebaCartas.pdf");
                cartaCumpleaniosBusiness.CrearCartasCumpleanios(cartas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void CrearEtiquetas(IEnumerable<CartaEntity> cartas)
        {
            try
            {
                CumpleaniosEtiquetaBusiness cumpleaniosEtiquetaBusiness = new CumpleaniosEtiquetaBusiness("C:\\Leo\\pruebaEtiquetas.pdf");
                cumpleaniosEtiquetaBusiness.CrearEtiquetas(cartas);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
