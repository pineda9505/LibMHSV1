using ClosedXML.Excel;
using Conexion;
using EventLogs;
using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
namespace LibMHSV1
{
    public class F07VtsContr : Declaracion
    {
        IConexion cn;
        ISQLPack sq;
        private byte Periodo;
        private int Ejercicio;
        public string NumDoc { get; set; }
        public int NumContInt { get; set; }
        public string NitNrc { get; set; }
        public double VtaExent { get; set; }
        public double VtaNoSuj { get; set; }
        public double VtaGrav { get; set; }
        public double IvaDeb { get; set; }
        public double VtaTer { get; set; }
        public double IvaDebTer { get; set; }
        public double TotVtas { get; set; }
        public string Dui { get; set; }

        public byte Anexo { get; } = 1;
        private  string expNit = @"^\d{4}-?\d{6}-?\d{3}-?\d{1}$";
        private  string expDui = @"^\d{8}-?\d{1}$";
        public List<F07VtsContr> Lista { get; set; }

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
                        hpis.Cell("D" + (i + 1)).Value = f.Resolucion;
                        hpis.Cell("E" + (i + 1)).Value = f.Serie;
                        hpis.Cell("F" + (i + 1)).Value = f.NumDoc;
                        hpis.Cell("G" + (i + 1)).Value = f.NumContInt;
                        if(f.Nit.Trim().Equals(""))
                        {
                            if (DocIdent.NRC.Validar(f.Nrc))
                            {
                                hpis.Cell("H" + (i + 1)).Value = f.Nrc;
                            }
                            else
                            {
                                hpis.Cell("H" + (i + 1)).Value = f.Nrc;
                            }
                        }
                        else
                        {
                            if(Regex.IsMatch(f.Nit,expNit))
                            {
                                if (DocIdent.NIT.Validar(f.Nit))
                                {
                                    hpis.Cell("H" + (i + 1)).Value = f.Nit;
                                }
                                else
                                {
                                    hpis.Cell("H" + (i + 1)).Value = f.Nrc;
                                }
                            }
                            else if (Regex.IsMatch(f.Nit, expDui))
                            {
                                if (DocIdent.DUI.Validar(f.Nit))
                                {
                                    hpis.Cell("H" + (i + 1)).Value = f.Nit;
                                }
                                else
                                {
                                    hpis.Cell("H" + (i + 1)).Value = f.Nit;
                                }
                            }
                            else
                            {
                                hpis.Cell("H" + (i + 1)).Value = f.Nit + "-Revisar";
                            }
                        }
                        hpis.Cell("I" + (i + 1)).Value = f.Nombre;
                        hpis.Cell("J" + (i + 1)).Value = f.VtaExent.ToString("####0.00");
                        hpis.Cell("K" + (i + 1)).Value = f.VtaNoSuj.ToString("####0.00");
                        hpis.Cell("L" + (i + 1)).Value = f.VtaGrav.ToString("####0.00");
                        hpis.Cell("M" + (i + 1)).Value = f.IvaDeb.ToString("####0.00");
                        hpis.Cell("N" + (i + 1)).Value = f.VtaTer.ToString("####0.00");
                        hpis.Cell("O" + (i + 1)).Value = f.IvaDebTer.ToString("####0.00");
                        hpis.Cell("P" + (i + 1)).Value = f.TotVtas.ToString("####0.00");
                        if (!f.DUI.Trim().Equals(""))
                        {
                            if (DocIdent.DUI.Validar(f.DUI))
                            {
                                hpis.Cell("Q" + (i + 1)).Value = f.DUI;
                            }
                            else
                            {
                                hpis.Cell("Q" + (i + 1)).Value = f.DUI;
                            }
                        }
                        else
                        {
                            hpis.Cell("Q" + (i + 1)).Value = f.DUI + "";
                        }
                        hpis.Cell("R" + (i + 1)).Value = f.Anexo;
                        i++;
                    }
                    hpis.ColumnsUsed().AdjustToContents();
                    try
                    {
                        vtsCons.SaveAs(ruta);
                    }
                    catch (Exception ex)
                    {
                        Registros.Escribir("Libros MH/ Contribuyentes", ex.Message);
                    }
                }
                return File.Exists(ruta);
            }
        }

        public F07VtsContr(byte periodo, int ejercicio)
        {
            cn = new ConexionSql();
            if (cn.Conecto())
            {
                try
                {
                    sq = new SQLPack();
                    sq.ProcedimientoAlmacenado(cn.BdFact + ".dbo.sp_LblMhHCtrbts", cn.getSqlcon());
                    sq.Parametros = new SqlParameter[2];
                    sq.Parametros[0] = new SqlParameter("@periodo", periodo)
                    {
                        SqlDbType = SqlDbType.SmallInt
                    };
                    sq.Parametros[1] = new SqlParameter("@ejercicio", ejercicio)
                    {
                        SqlDbType = System.Data.SqlDbType.Int
                    };
                    sq.AsignarParametros();
                    cn.Conectar();
                    if (sq.RetornarFilas().HasRows)
                    {
                        Lista = new List<F07VtsContr>();
                        while (sq.Dr.Read())
                        {
                            Lista.Add(new F07VtsContr(Convert.ToDateTime(sq.Dr["FECHA EMISION DEL DOCUMENTO"])
                                                                        , Convert.ToByte(sq.Dr["CLASE DE DOCUMENTO"])
                                                                        , sq.Dr["TIPO DOCUMENTO"].ToString()
                                                                        , sq.Dr["NRC"].ToString()
                                                                        , sq.Dr["NIT"].ToString()
                                                                        , sq.Dr["DUI"].ToString()
                                                                        , sq.Dr["NOMBRE"].ToString()
                                                                        , sq.Dr["RESOLUCION"].ToString()
                                                                        , sq.Dr["SERIE DEL DOCUMENTO"].ToString()
                                                                        , sq.Dr["NUMERO DE DOCUMENTO"].ToString()
                                                                        , Convert.ToInt32(sq.Dr["CORRELATIVO INTERNO"])
                                                                        , Convert.ToDouble(sq.Dr["VENTAS EXENTAS"])
                                                                        , Convert.ToDouble(sq.Dr["VENTAS NO SUJETAS"])
                                                                        , Convert.ToDouble(sq.Dr["VENTAS GRAVADAS LOCALES"])
                                                                        , Convert.ToDouble(sq.Dr["IVA"])
                                                                        , Convert.ToDouble(sq.Dr["VENTAS A CUENTAS DE TERCEROS"])
                                                                        , Convert.ToDouble(sq.Dr["DEBITO FISCAL A CUENTA DE TERCEROS"])
                                                                        , Convert.ToDouble(sq.Dr["TOTAL VENTAS"])
                                                                        , Convert.ToByte(sq.Dr["ANEXO"])));
                        }
                    }
                   
                }
                catch (Exception ex)
                {
                    Registros.Escribir("Libros MH/ Contribuyente", ex.Message);
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

        public F07VtsContr(DateTime fecha, byte claseDoc, string tipodoc, string nrc, string nit, string dUI, string nombre, string resolucion, string serie, string numDoc, int numContInt, double vtaExent, double vtaNoSuj, double vtaGrav, double ivaDeb, double vtaTer, double ivaDebTer, double totVtas, byte anexo)
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
            NumDoc = numDoc;
            NumContInt = numContInt;
            VtaExent = vtaExent;
            VtaNoSuj = vtaNoSuj;
            VtaGrav = vtaGrav;
            IvaDeb = ivaDeb;
            VtaTer = vtaTer;
            IvaDebTer = ivaDebTer;
            TotVtas = totVtas;
            Anexo = anexo;
        }
    }
}
