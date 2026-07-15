using ClosedXML.Excel;
using Conexion;
using EventLogs;
using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.IO;

namespace LibMHSV1
{
    public class F07Percepciones:Declaracion
    {
        IConexion cn;
        ISQLPack sq;
        public List<F07Percepciones> Lista { get; set; }
        private byte Periodo;
        private int Ejercicio;


        public string NumDoc { get; set; }
        public double SubTot { get; set; }
        public double Retencion { get; set; }

        public byte Anexo { get; }
        public F07Percepciones(string nit, DateTime fecha, string tipoDoc, string serie, string numDoc, double subTot, double retencion, string dui, byte anexo)
        {
            Nit = nit;
            Fecha = fecha;
            Tipodoc = tipoDoc;
            Serie = serie;
            NumDoc = numDoc;
            SubTot = subTot;
            Retencion = retencion;
            DUI = dui;
            Anexo = anexo;
        }

        public F07Percepciones(byte periodo, int ejercicio)
        {
            cn = new ConexionSql();
            if (cn.Conecto())
            {
                try
                {
                    sq = new SQLPack();
                    sq.ProcedimientoAlmacenado("sp_F07PercepcionesMostrar", cn.getSqlcon());
                    sq.Parametros = new SqlParameter[2];
                    sq.Parametros[0] = new SqlParameter("@periodo", periodo)
                    {
                        SqlDbType = SqlDbType.SmallInt
                    };
                    sq.Parametros[1] = new Microsoft.Data.SqlClient.SqlParameter("@ejercicio", ejercicio)
                    {
                        SqlDbType = System.Data.SqlDbType.Int
                    };
                    sq.AsignarParametros();
                    cn.Conectar();
                    if (sq.RetornarFilas().HasRows)
                    {
                        Lista = new List<F07Percepciones>();
                        while (sq.Dr.Read())
                        {
                            Lista.Add(new F07Percepciones(sq.Dr["NIT"].ToString()
                                                         , Convert.ToDateTime(sq.Dr["FECHA EMISION DEL DOCUMENTO"])
                                                         , sq.Dr["TIPO DOCUMENTO"].ToString()
                                                         , sq.Dr["SERIE DEL DOCUMENTO"].ToString()
                                                         , sq.Dr["NUMERO DE DOCUMENTO"].ToString()
                                                         , Convert.ToDouble(sq.Dr["SUBTOTAL"])
                                                         , Convert.ToDouble(sq.Dr["PERCEPCION"])
                                                         , sq.Dr["DUI"].ToString()
                                                         , Convert.ToByte(sq.Dr["ANEXO"])
                                                         ));
                        }
                    }

                }
                catch (Exception ex)
                {
                    Registros.Escribir("Libros MH / Percepciones", ex.Message);
                }
                finally
                {
                    cn.Desconectar();
                    sq.DrClose();
                }
            }
            Periodo = periodo;
            Ejercicio = ejercicio;
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

                        if (f.Nit.Trim().Equals(""))
                        {
                            if (!(f.Nrc == null))
                            {
                                if (DocIdent.NRC.Validar(f.Nrc))
                                {
                                    hpis.Cell("A" + (i + 1)).Value = f.Nrc;
                                }
                                else
                                {
                                    hpis.Cell("A" + (i + 1)).Value = f.Nrc + "R";
                                }
                            }
                            else
                            {
                                hpis.Cell("A" + (i + 1)).Value = "";
                            }

                        }
                        else
                        {
                            if (f.Nit.Replace("-", "").Length == 9)
                            {
                                if (DocIdent.DUI.Validar(f.Nit))
                                {
                                    hpis.Cell("A" + (i + 1)).Value = f.Nit;
                                }
                                else
                                {
                                    hpis.Cell("A" + (i + 1)).Value = f.Nrc + "R";
                                }
                            }
                            else if (f.Nit.Replace("-", "").Length == 14)
                            {


                                if (DocIdent.NIT.Validar(f.Nit))
                                {
                                    hpis.Cell("A" + (i + 1)).Value = f.Nit;
                                }
                                else
                                {
                                    hpis.Cell("A" + (i + 1)).Value = f.Nrc + "R";
                                }
                            }
                        }
                        hpis.Cell("B" + (i + 1)).Value = f.Fecha.ToString("dd/MM/yyyy");
                        hpis.Cell("C" + (i + 1)).Value = f.Tipodoc.ToString();
                        hpis.Cell("D" + (i + 1)).Value = f.Serie;
                        hpis.Cell("E" + (i + 1)).Value = f.NumDoc;
                        hpis.Cell("F" + (i + 1)).Value = f.SubTot.ToString("####0.00");
                        hpis.Cell("G" + (i + 1)).Value = f.Retencion.ToString("####0.00");
                        if (!f.DUI.Trim().Equals(""))
                        {
                            if (DocIdent.DUI.Validar(f.DUI))
                            {
                                hpis.Cell("H" + (i + 1)).Value = (f.Nit.Equals("")) ? f.DUI : "";
                            }
                            else
                            {
                                hpis.Cell("H" + (i + 1)).Value = (f.Nit.Equals("")) ? f.DUI + "R" : "";
                            }
                        }
                        else
                        {

                        }
                        hpis.Cell("I" + (i + 1)).Value = f.Anexo;
                        i++;
                    }
                    hpis.ColumnsUsed().AdjustToContents();
                    try
                    {
                        vtsCons.SaveAs(ruta);
                    }
                    catch (Exception ex)
                    {
                        Registros.Escribir("Libros MH / Percepciones", ex.Message);
                    }
                }
                return File.Exists(ruta);
            }
        }
    }
}
