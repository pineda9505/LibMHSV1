using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.IO;
using System.Data;
using Microsoft.Data.SqlClient;
using EventLogs;
using Conexion;
namespace LibMHSV1
{
    public class F07VtsConsum : Declaracion
    {
        IConexion cn;
        ISQLPack sq;
        private byte Periodo;
        private int ejercicio;
        public List<F07VtsConsum> Lista { get; set; }
        public string DocDel { get; set; }
        public string DocAl { get; set; }
        public string NumCtlIntDel { get; set; }
        public string NumCtlIntAl { get; set; }
        public string NumMaqReg { get; set; }
        public double Exenta { get; set; }
        public double IntExtNSProp { get; set; }
        public double NSujeta { get; set; }
        public double GravLcl { get; set; }
        public double ExportAraCentAm { get; set; }
        public double ExportFraAraCentam { get; set; }
        public double ExportServ { get; set; }
        public double ZnFranc { get; set; }
        public double VtaTerc { get; set; }
        public double Total { get; set; }
        public byte Anexo { get; }

        public bool ExportarAExcel(string ruta)
        {
            using (XLWorkbook vtsCons = new XLWorkbook())
            {
                var hpis = vtsCons.Worksheets.Add("FA-" + Periodo.ToString() + "-" + ejercicio.ToString());
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
                        hpis.Cell("F" + (i + 1)).Value = f.DocDel;
                        hpis.Cell("G" + (i + 1)).Value = f.DocAl;
                        hpis.Cell("H" + (i + 1)).Value = f.NumCtlIntDel;
                        hpis.Cell("I" + (i + 1)).Value = f.NumCtlIntAl;
                        hpis.Cell("J" + (i + 1)).Value = f.NumMaqReg;
                        hpis.Cell("K" + (i + 1)).Value = f.Exenta.ToString("####0.00");
                        hpis.Cell("L" + (i + 1)).Value = f.IntExtNSProp.ToString("####0.00");
                        hpis.Cell("M" + (i + 1)).Value = f.NSujeta.ToString("####0.00");
                        hpis.Cell("N" + (i + 1)).Value = f.GravLcl.ToString("####0.00");
                        hpis.Cell("O" + (i + 1)).Value = f.ExportAraCentAm.ToString("####0.00");
                        hpis.Cell("P" + (i + 1)).Value = f.ExportFraAraCentam.ToString("####0.00");
                        hpis.Cell("Q" + (i + 1)).Value = f.ExportServ.ToString("####0.00");
                        hpis.Cell("R" + (i + 1)).Value = f.ZnFranc.ToString("####0.00");
                        hpis.Cell("S" + (i + 1)).Value = f.VtaTerc.ToString("####0.00");
                        hpis.Cell("T" + (i + 1)).Value = f.Total.ToString("####0.00");
                        hpis.Cell("U" + (i + 1)).Value = f.Anexo.ToString("###");
                        i++;
                    }
                    hpis.ColumnsUsed().AdjustToContents();
                    try
                    {
                        vtsCons.SaveAs(ruta);
                    }
                    catch (Exception ex)
                    {
                        Registros.Escribir("Libros MH/ Consumidor Final", ex.Message);
                    }
                }
                return File.Exists(ruta);
            }
        }

        public bool GuardarComoCSV(string ruta)
        {
            using (XLWorkbook vtsCons = new XLWorkbook())
            {
                var hpis = vtsCons.Worksheets.Add("FA-" + Periodo.ToString() + "-" + ejercicio.ToString());
                int i = 0;
                foreach (var f in Lista)
                {
                    hpis.Cell("A" + (i + 1)).Value = f.Fecha.ToString("dd/MM/yyyy");
                    hpis.Cell("B" + (i + 1)).Value = f.ClaseDoc;
                    hpis.Cell("C" + (i + 1)).Value = f.Tipodoc.ToString();
                    hpis.Cell("D" + (i + 1)).Value = f.Resolucion;
                    hpis.Cell("E" + (i + 1)).Value = f.Serie;
                    hpis.Cell("F" + (i + 1)).Value = f.DocDel;
                    hpis.Cell("G" + (i + 1)).Value = f.DocAl;
                    hpis.Cell("H" + (i + 1)).Value = f.NumCtlIntDel;
                    hpis.Cell("I" + (i + 1)).Value = f.NumCtlIntAl;
                    hpis.Cell("J" + (i + 1)).Value = f.NumMaqReg;
                    hpis.Cell("K" + (i + 1)).Value = f.Exenta.ToString("####0.00");
                    hpis.Cell("L" + (i + 1)).Value = f.IntExtNSProp.ToString("####0.00");
                    hpis.Cell("M" + (i + 1)).Value = f.NSujeta.ToString("####0.00");
                    hpis.Cell("N" + (i + 1)).Value = f.GravLcl.ToString("####0.00");
                    hpis.Cell("O" + (i + 1)).Value = f.ExportAraCentAm.ToString("####0.00");
                    hpis.Cell("O" + (i + 1)).Value = f.ExportFraAraCentam.ToString("####0.00");
                    hpis.Cell("O" + (i + 1)).Value = f.ExportServ.ToString("####0.00");
                    hpis.Cell("O" + (i + 1)).Value = f.ZnFranc.ToString("####0.00");
                    hpis.Cell("O" + (i + 1)).Value = f.VtaTerc.ToString("####0.00");
                    hpis.Cell("O" + (i + 1)).Value = f.Total.ToString("####0.00");
                    hpis.Cell("O" + (i + 1)).Value = f.Anexo.ToString("###");
                    i++;
                }
                hpis.ColumnsUsed().AdjustToContents();

                try
                {
                    vtsCons.SaveAs(ruta);
                }
                catch (Exception ex)
                {

                }
                return File.Exists(ruta);
            }
        }


        public F07VtsConsum(int periodo, int ejercicio)
        {
            cn = new ConexionSql();
            if (cn.Conecto())
            {
                try
                {
                    sq = new Conexion.SQLPack();
                    sq.ProcedimientoAlmacenado(cn.BdFact + ".dbo.sp_LblMHCnsFin", cn.getSqlcon());
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
                        Lista = new List<F07VtsConsum>();
                        while (sq.Dr.Read())
                        {
                            Lista.Add(new F07VtsConsum(Convert.ToDateTime(sq.Dr["FECHA EMISION DEL DOCUMENTO"])
                                                        , Convert.ToByte(sq.Dr["CLASE DE DOCUMENTO"])
                                                        , sq.Dr["TIPO DOCUMENTO"].ToString()
                                                        , sq.Dr["RESOLUCION"].ToString()
                                                        , sq.Dr["SERIE DEL DOCUMENTO"].ToString()
                                                        , sq.Dr["NUMERO DE DOCUMENTO DEL"].ToString()
                                                        , sq.Dr["NUMERO DE DOCUMENTO AL"].ToString()
                                                        , sq.Dr["CORRELATIVO INTERNO DEL"].ToString()
                                                        , sq.Dr["CORRELATIVO INTERNO AL"].ToString()
                                                        , sq.Dr["NUMERO DE MAQUINA REGISTRADORA"].ToString()
                                                        , Convert.ToDouble(sq.Dr["VENTAS EXENTAS"])
                                                        , Convert.ToDouble(sq.Dr["VENTAS INTERNA EXENTA NO SUJETA A PROPORCIONALIDAD"])
                                                        , Convert.ToDouble(sq.Dr["VENTA NO SUJETA"])
                                                        , Convert.ToDouble(sq.Dr["VENTAS GRAVADAS LOCALES"])
                                                        , Convert.ToDouble(sq.Dr["EXPORTACIONES DENTRO DEL AREA CENTROAMERICANA"])
                                                        , Convert.ToDouble(sq.Dr["EXPORTACIONES FUERA DEL AREA CENTROAMERICANA"])
                                                        , Convert.ToDouble(sq.Dr["EXPORTACION DE SERVICIOS"])
                                                        , Convert.ToDouble(sq.Dr["VENTAS A ZONAS FRANCAS Y DPA (TASA CERO)"])
                                                        , Convert.ToDouble(sq.Dr["VENTA A CUENTA DE TERCEROS NO DOMICILIADOS"])
                                                        , Convert.ToDouble(sq.Dr["TOTAL VENTAS"])
                                                        , Convert.ToByte(sq.Dr["ANEXO"])));
                        }
                    }

                }
                catch (Exception ex)
                {
                    Registros.Escribir("Libros MH/ Consumidor Final", ex.Message);
                }
                finally
                {
                    cn.Desconectar();
                    sq.DrClose();
                }
            }
        }

        public F07VtsConsum(DateTime fecha, byte claseDoc, string tipodoc, string resolucion, string serie, string docDel, string docAl, string numCtlIntDel, string numCtlIntAl, string numMaqReg, double exenta, double intExtNSProp, double nSujeta, double gravLcl, double exportAraCentAm, double exportFraAraCentam, double exportServ, double znFranc, double vtaTerc, double total, byte anexo)
        {
            Fecha = fecha;
            ClaseDoc = claseDoc;
            Tipodoc = tipodoc;
            Resolucion = resolucion;
            Serie = serie;
            DocDel = docDel;
            DocAl = docAl;
            NumCtlIntDel = numCtlIntDel;
            NumCtlIntAl = numCtlIntAl;
            NumMaqReg = numMaqReg;
            Exenta = exenta;
            IntExtNSProp = intExtNSProp;
            NSujeta = nSujeta;
            GravLcl = gravLcl;
            ExportAraCentAm = exportAraCentAm;
            ExportFraAraCentam = exportFraAraCentam;
            ExportServ = exportServ;
            ZnFranc = znFranc;
            VtaTerc = vtaTerc;
            Total = total;
            Anexo = anexo;
        }
    }
}
