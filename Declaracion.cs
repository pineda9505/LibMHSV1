using System;

namespace LibMHSV1
{
    public abstract class Declaracion
    {
        public DateTime Fecha { get; set; }
        public byte ClaseDoc { get; set; }
        public string Tipodoc { get; set; }
        public string Nrc { get; set; }
        public string Nit { get; set; }
        public string DUI { get; set; }
        public string Nombre { get; set; }
        public string Resolucion { get; set; }
        public string Serie { get; set; }
        /*
        protected Declaracion(DateTime fecha, byte claseDoc, string tipodoc, string nrc, string nit, string dUI, string nombre, string resolucion, string serie)
        {
            Fecha = fecha;
            ClaseDoc = claseDoc;
            Tipodoc = tipodoc;
            Nrc = nrc;
            Nit = nit;
            DUI = dUI;
            Nombre = nombre;
            Resolucion = resolucion;
            Serie = serie;
        }
        */
    }
}
