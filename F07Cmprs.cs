using ClosedXML.Excel;
using Conexion;
using EventLogs;
using System;
using System.Collections.Generic;
using System.IO;

namespace LibMHSV1
{
    public class F07Cmprs : Declaracion
    {

        public string NDoc { get; set; }
        public double ExtNSuj { get; set; }
        public double InterExtNoSuj { get; set; }
        public double ImpExtNoSuj { get; set; }
        public double ComIntGrav { get; set; }
        public double IntGravBienes { get; set; }
        public double ImpGravBienes { get; set; }
        public double ImpGravServicios { get; set; }
        public double CredFis { get; set; }
        public double TotComp { get; set; }
        public byte Anexo { get; }

        Conexion.IConexion cn;
        Conexion.ISQLPack sq;
        private byte Periodo;
        private int Ejercicio;
        public List<F07Cmprs> Lista { get; set; }
        public F07Cmprs(byte periodo, int ejercicio)
        {
            try
            {
                cn = new ConexionSql();
                if (cn.Conecto())
                {
                    sq = new Conexion.SQLPack();
                    sq.ProcedimientoAlmacenado("sp_LblMhComp", cn.getSqlcon());
                    sq.Parametros = new Microsoft.Data.SqlClient.SqlParameter[2];
                    sq.Parametros[0] = new Microsoft.Data.SqlClient.SqlParameter("periodo", periodo)
                    {
                        SqlDbType = System.Data.SqlDbType.SmallInt
                    };
                    sq.Parametros[1] = new Microsoft.Data.SqlClient.SqlParameter("ejercicio", ejercicio)
                    {
                        SqlDbType = System.Data.SqlDbType.Int
                    };
                    sq.AsignarParametros();
                    cn.Conectar();
                    if (sq.RetornarFilas().HasRows)
                    {
                        Lista = new List<F07Cmprs>();
                        while (sq.Dr.Read())
                        {
                            Lista.Add(
                                new F07Cmprs(
                                                fecha: Convert.ToDateTime(sq.Dr["Fecha de emision"])
                                                , claseDoc: Convert.ToByte(sq.Dr["Clase de documento"])
                                                , Convert.ToInt16(sq.Dr["Tipo documento"]).ToString("00")
                                                , sq.Dr["NRC"].ToString()
                                                , sq.Dr["NIT"].ToString()
                                                , sq.Dr["DUI"].ToString()
                                                , sq.Dr["Nombre"].ToString()
                                                , ""
                                                , sq.Dr["serie"].ToString()
                                                , sq.Dr["Numero de documento"].ToString()
                                                , Convert.ToDouble(sq.Dr["Compras internas exentas no sujetas"])
                                                , Convert.ToDouble(sq.Dr["Internaciones exentas no sujetas"])
                                                , Convert.ToDouble(sq.Dr["Importaciones exentas no sujetas"])
                                                , Convert.ToDouble(sq.Dr["Compras internas gravadas"])
                                                , Convert.ToDouble(sq.Dr["Internaciones gravadas de bienes"])
                                                , Convert.ToDouble(sq.Dr["Importaciones gravadas de bienes"])
                                                , Convert.ToDouble(sq.Dr["Importaciones gravadas de servicios"])
                                                , Convert.ToDouble(sq.Dr["Crédito fiscal"])
                                                , Convert.ToDouble(sq.Dr["Total compras"])
                                                , 3
                                            )
                                );
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Registros.Escribir("Libros MH/ Compras", ex.Message);
            }
            finally
            {
                cn.Desconectar();
            }
            Periodo = periodo;
            Ejercicio = ejercicio;
        }

        public F07Cmprs(DateTime fecha, byte claseDoc, string tipodoc, string nrc, string nit, string dUI, string nombre, string resolucion, string serie, string nDoc, double extNSuj, double interExtNoSuj, double impExtNoSuj, double comIntGrav, double intGravBienes, double impGravBienes, double impGravServicios, double credFis, double totComp, byte anexo)
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
            NDoc = nDoc;
            ExtNSuj = extNSuj;
            InterExtNoSuj = interExtNoSuj;
            ImpExtNoSuj = impExtNoSuj;
            ComIntGrav = comIntGrav;
            IntGravBienes = intGravBienes;
            ImpGravBienes = impGravBienes;
            ImpGravServicios = impGravServicios;
            CredFis = credFis;
            TotComp = totComp;
            Anexo = anexo;
        }

        public bool ExportarAExcel(string ruta)
        {
            using (XLWorkbook vtsCons = new XLWorkbook())
            {
                var hpis = vtsCons.Worksheets.Add("CF-" + Periodo.ToString() + "-" + Ejercicio.ToString());
                int i = 0;
                if (!(Lista == null))
                {
                    foreach (var f in Lista)
                    {
                        hpis.Cell("A" + (i + 1)).Value = f.Fecha.ToString("dd/MM/yyyy");
                        hpis.Cell("B" + (i + 1)).Value = f.ClaseDoc;
                        hpis.Cell("C" + (i + 1)).Value = f.Tipodoc.ToString();
                        hpis.Cell("D" + (i + 1)).Value = (f.ClaseDoc == 4)?f.Serie: f.NDoc;
                        if (f.Nit.Equals(""))
                        {
                            if (DocIdent.NRC.Validar(f.Nrc))
                            {
                                hpis.Cell("E" + (i + 1)).Value = f.Nrc;
                            }
                            else
                            {
                                hpis.Cell("E" + (i + 1)).Value = f.Nrc + "R";
                            }
                        }
                        else
                        {
                            if (f.Nit.Length > 0)
                            {
                                if (DocIdent.NIT.Validar(f.Nit))
                                {
                                    hpis.Cell("E" + (i + 1)).Value = f.Nit;
                                }
                                else
                                {
                                    hpis.Cell("E" + (i + 1)).Value = f.Nrc + "R";
                                }
                            }
                            
                        }
                        hpis.Cell("F" + (i + 1)).Value = f.Nombre;
                        hpis.Cell("G" + (i + 1)).Value = f.ExtNSuj.ToString("###0.00");
                        hpis.Cell("H" + (i + 1)).Value = f.InterExtNoSuj.ToString("###0.00");
                        hpis.Cell("I" + (i + 1)).Value = f.ImpExtNoSuj.ToString("####0.00"); ;
                        hpis.Cell("J" + (i + 1)).Value = f.ComIntGrav.ToString("####0.00");
                        hpis.Cell("K" + (i + 1)).Value = f.IntGravBienes.ToString("####0.00");
                        hpis.Cell("L" + (i + 1)).Value = f.ImpGravBienes.ToString("####0.00");
                        hpis.Cell("M" + (i + 1)).Value = f.ImpGravServicios.ToString("####0.00");
                        hpis.Cell("N" + (i + 1)).Value = f.CredFis.ToString("####0.00");
                        hpis.Cell("O" + (i + 1)).Value = f.TotComp.ToString("####0.00");
                        if (!f.DUI.Trim().Equals(""))
                        {
                            if (DocIdent.DUI.Validar(f.DUI))
                            {
                                hpis.Cell("P" + (i + 1)).Value = f.DUI;
                            }
                            else
                            {
                                hpis.Cell("P" + (i + 1)).Value = f.DUI + "R";
                            }
                        }
                        else
                        {
                            hpis.Cell("P" + (i + 1)).Value = f.DUI + "";
                        }
                        hpis.Cell("Q" + (i + 1)).Value = f.Anexo;
                        i++;
                    }
                    hpis.ColumnsUsed().AdjustToContents();
                    try
                    {
                        vtsCons.SaveAs(ruta);
                    }
                    catch (Exception ex)
                    {
                        Registros.Escribir("Libros MH/ Compras", ex.Message);
                    }
                }
                return File.Exists(ruta);
            }
        }
    }
}
