using OfficeOpenXml.Style;
using OfficeOpenXml;
using SLSEARAPI.DataLayer;
using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace SLSEARAPI.Controllers
{
    public class PlanAsistenciaTecController : ApiController
    {
        PlanAsistenciaTecDL asistenciaTecDL;
        public PlanAsistenciaTecController()
        {
            asistenciaTecDL = new PlanAsistenciaTecDL();
        }

        [HttpPost]
        [ActionName("ListarPlanAsistenciaTec")]
        public List<PlanAsistenciaTec> ListarPlanAsistenciaTec(PlanAsistenciaTec planAsistenciaTec)
        {
            try
            {
                return asistenciaTecDL.ListarPlanAsistenciaTec(planAsistenciaTec);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarPlanAsistenciaTec2")]
        public List<PlanAsistenciaTec> ListarPlanAsistenciaTec2(PlanAsistenciaTec planAsistenciaTec)
        {
            try
            {
                return asistenciaTecDL.ListarPlanAsistenciaTec2(planAsistenciaTec);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarPlanAsistenciaTecDet")]
        public List<PlanAsistenciaTecDet> ListarPlanAsistenciaTecDet(PlanAsistenciaTecDet planAsistenciaTecDet)
        {
            try
            {
                return asistenciaTecDL.ListarPlanAsistenciaTecDet(planAsistenciaTecDet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarPlanAsistenciaTec")]
        public PlanAsistenciaTec InsertarPlanAsistenciaTec(PlanAsistenciaTec planAsistenciaTec)
        {
            try
            {
                return asistenciaTecDL.InsertarPlanAsistenciaTec(planAsistenciaTec);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarPlanAsistenciaTecDet")]
        public PlanAsistenciaTecDet InsertarPlanAsistenciaTecDet(PlanAsistenciaTecDet planAsistenciaTecDet)
        {
            try
            {
                return asistenciaTecDL.InsertarPlanAsistenciaTecDet(planAsistenciaTecDet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ExportarPlanAsistenciaTec")]
        public HttpResponseMessage ExportarPlanAsistenciaTec(Actividad actividad)
        {
            try
            {
                String NombreReporte = "ExportarPlanAsistenciaTec";

                using (var excelPackage = new ExcelPackage())
                {
                    excelPackage.Workbook.Properties.Author = NombreReporte;
                    excelPackage.Workbook.Properties.Title = NombreReporte;
                    var _genericSheet = excelPackage.Workbook.Worksheets.Add(NombreReporte);
                    _genericSheet.View.ShowGridLines = false;
                    _genericSheet.View.ZoomScale = 100;
                    _genericSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                    _genericSheet.PrinterSettings.FitToPage = true;
                    _genericSheet.PrinterSettings.Orientation = eOrientation.Landscape;
                    _genericSheet.View.PageBreakView = true;


                    //List<string> _cabecera = new List<string>();

                    //_cabecera.Add("Nro");
                    //_cabecera.Add("Actividades");
                    //_cabecera.Add("Descripcion");
                    //_cabecera.Add("Unidad Medida");
                    //_cabecera.Add("Meta");

                    //int finish = pintarcabeceras(_cabecera, _genericSheet, "CRONOGRAMA DETALLADO");

                    //_genericSheet.Cells["A3:D3"].Value = "3.1 CRONOGRAMA DE EJECUCIÓN";
                    //_genericSheet.Cells["A3:D3"].Merge = true;
                    //_genericSheet.Cells["A3:D3"].Style.Locked = true;
                    //System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#72aea5");
                    //_genericSheet.Cells["A3:D3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Cells["A3:D3"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //pintar componente 1

                    //List<Componente> componentes = cronogramaDL.ListarComponentes(cronograma);
                    DataTable SearTab = asistenciaTecDL.SP_Listar_PlanAsistenciaTecnica_Rpt(actividad);
                    int rowIndexComp = 1;
                    int colcomp = 2;

                    if (SearTab.Rows.Count > 0)
                    {
                        for (int i = 0; i < SearTab.Rows.Count; i++)
                        {
                            _texto_row(_genericSheet, rowIndexComp, colcomp, "", "#72AEA5");
                            //colcomp =+ 7;
                            _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                            rowIndexComp++;
                            //colcomp =- 7;


                            _texto_row(_genericSheet, rowIndexComp, colcomp, "PLAN DE ASISTENCIA TÉCNICA", "#72AEA5");
                            _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                            _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.Font.Bold = true;
                            _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            rowIndexComp++;
                            //colcomp = -7;

                            _texto_row(_genericSheet, rowIndexComp, colcomp, "", "#72AEA5");
                            _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                            rowIndexComp++;
                            //colcomp = -7;

                            // Cabecera de Plan Capacitación
                            _texto_row(_genericSheet, rowIndexComp, colcomp, "Nombre del SEAR", "#ffffff");
                            _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                            _texto_row(_genericSheet, rowIndexComp, colcomp + 2, SearTab.Rows[i][0], "#ffffff");
                            _genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                            //colcomp = 2;
                            rowIndexComp++;
                            //colcomp--;
                            _texto_row(_genericSheet, rowIndexComp, colcomp, "Nombre de extensionista", "#ffffff");
                            _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                            _texto_row(_genericSheet, rowIndexComp, colcomp + 2, SearTab.Rows[i][1], "#ffffff");
                            _genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                            //colcomp = 2;
                            rowIndexComp++;
                            //colcomp--;
                            _texto_row(_genericSheet, rowIndexComp, colcomp, "Agencia Agraria", "#ffffff");
                            _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                            _texto_row(_genericSheet, rowIndexComp, colcomp + 2, SearTab.Rows[i][3], "#ffffff");
                            _genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                            //colcomp = 2;
                            rowIndexComp++;
                            //colcomp--;
                            _texto_row(_genericSheet, rowIndexComp, colcomp, "Dirección Zonal", "#ffffff");
                            _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                            _texto_row(_genericSheet, rowIndexComp, colcomp + 2, SearTab.Rows[i][4], "#ffffff");
                            _genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                            //colcomp = 2;
                            rowIndexComp++;
                            //colcomp--;
                            _texto_row(_genericSheet, rowIndexComp, colcomp, SearTab.Rows[i][8], "#72AEA5");
                            _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                            _texto_row(_genericSheet, rowIndexComp, colcomp + 2, SearTab.Rows[i][9], "#72AEA5");
                            _genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;


                            /*****************************************************************************************/
                            PlanAsistenciaTec planAsistenciaTec = new PlanAsistenciaTec();
                            planAsistenciaTec.iCodActividad = Convert.ToInt32(SearTab.Rows[i][7]);
                            DataTable Modulotab = asistenciaTecDL.SP_Listar_PlanAsistenciaTecnica_Rpt2(planAsistenciaTec);
                            //rowIndexComp = 1;
                            //colcomp = 2;
                            for (int j = 0; j < Modulotab.Rows.Count; j++)
                            {
                                //rowIndexComp++;
                                //_texto_row(_genericSheet, rowIndexComp, colcomp, "Modulo o tema", "#ffffff");
                                //_genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                                //_texto_row(_genericSheet, rowIndexComp, colcomp + 2, Modulotab.Rows[j][0], "#E2EFDA");
                                //_genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;

                                rowIndexComp++;
                                _texto_row(_genericSheet, rowIndexComp, colcomp, "Objetivo de la sesión", "#ffffff");
                                _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 2, Modulotab.Rows[j][1], "#E2EFDA");
                                _genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;

                                //Salto de Linea
                                rowIndexComp++;

                                _texto_row(_genericSheet, rowIndexComp, colcomp, "Meta (productores)", "#ffffff");
                                _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 2, Modulotab.Rows[j][2], "#E2EFDA");
                                //_genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 3, "Beneficiarios", "#ffffff");
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 4, Modulotab.Rows[j][3], "#E2EFDA");
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 5, "Fecha", "#ffffff");
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 6, Modulotab.Rows[j][4], "#E2EFDA");

                                //Salto de Linea
                                rowIndexComp++;

                                _texto_row(_genericSheet, rowIndexComp, colcomp, "Duración Total (horas)", "#ffffff");
                                _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 2, Modulotab.Rows[j][7], "#E2EFDA");
                                //_genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 3, "Teoria", "#ffffff");
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 4, Modulotab.Rows[j][5], "#E2EFDA");
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 5, "Práctica", "#ffffff");
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 6, Modulotab.Rows[j][6], "#E2EFDA");

                                rowIndexComp++;

                                _texto_row(_genericSheet, rowIndexComp, colcomp, "Lugar de ejecución", "#ffffff");
                                _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 2, Modulotab.Rows[j][8], "#E2EFDA");
                                _genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;

                                rowIndexComp++;
                                _texto_row(_genericSheet, rowIndexComp, colcomp, "Plan de sesión", "#ffffff");
                                _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                                _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.Font.Bold = true;
                                _genericSheet.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                rowIndexComp++;
                                _texto_row(_genericSheet, rowIndexComp, colcomp, "Duración", "#ffffff");
                                //_texto_row(_genericSheet, rowIndexComp, colcomp + 1, "Tematica / Pasos", "#ffffff");
                                //_genericSheet.Cells[rowIndexComp, colcomp + 1, rowIndexComp, colcomp + 2].Merge = true;
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 1, "Descripción de la Metodología", "#ffffff");
                                _genericSheet.Cells[rowIndexComp, colcomp + 1, rowIndexComp, colcomp + 4].Merge = true;
                                _texto_row(_genericSheet, rowIndexComp, colcomp + 5, "Instrumentos", "#ffffff");
                                _genericSheet.Cells[rowIndexComp, colcomp + 5, rowIndexComp, colcomp + 6].Merge = true;

                                /*****************************************************************************************/
                                PlanAsistenciaTecDet planPlanAsistenciaTecDet = new PlanAsistenciaTecDet();
                                planPlanAsistenciaTecDet.iCodPlanAsistenciaTec = Convert.ToInt32(Modulotab.Rows[j][9]);
                                DataTable SessionModtab = asistenciaTecDL.SP_Listar_PlanAsistenciaTecnica_Rpt3(planPlanAsistenciaTecDet);
                                for (int k = 0; k < SessionModtab.Rows.Count; k++)
                                {
                                    rowIndexComp++;
                                    _texto_row(_genericSheet, rowIndexComp, colcomp, SessionModtab.Rows[k][0], "#E2EFDA");
                                    //_texto_row(_genericSheet, rowIndexComp, colcomp + 1, SessionModtab.Rows[k][1], "#E2EFDA");
                                    //_genericSheet.Cells[rowIndexComp, colcomp + 1, rowIndexComp, colcomp + 2].Merge = true;
                                    _texto_row(_genericSheet, rowIndexComp, colcomp + 1, SessionModtab.Rows[k][2], "#E2EFDA");
                                    _genericSheet.Cells[rowIndexComp, colcomp + 1, rowIndexComp, colcomp + 4].Merge = true;
                                    _texto_row(_genericSheet, rowIndexComp, colcomp + 5, SessionModtab.Rows[k][3], "#E2EFDA");
                                    _genericSheet.Cells[rowIndexComp, colcomp + 5, rowIndexComp, colcomp + 6].Merge = true;
                                }
                            }

                            rowIndexComp = 1;
                            colcomp = colcomp + 8;
                            
                        }
                        int totalmetas = 0;
                    }

                    var x = excelPackage.GetAsByteArray();
                    string nombre = NombreReporte + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".xlsx";
                    var stream = new MemoryStream(x);
                    var result = Request.CreateResponse(HttpStatusCode.OK);
                    result.Content = new StreamContent(stream);
                    result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

                    result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = nombre
                    };
                    return result;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        internal void _texto_row(ExcelWorksheet _sheet, int rowIndex, int col, object _text, string Color = "", string fontName = "Calibri")
        {
            _sheet.Cells[rowIndex, col].Value = _text;
            _sheet.Cells[rowIndex, col].Merge = true;
            //_sheet.Columns[col].Width = 3;
            _sheet.Cells[rowIndex, col].Style.Font.Name = fontName;
            _sheet.Cells[rowIndex, col].Style.Font.Size = 12;
            _sheet.Cells[rowIndex, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            _sheet.Cells[rowIndex, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[rowIndex, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[rowIndex, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[rowIndex, col].Style.WrapText = false;
            _sheet.Columns[col].AutoFit();

            if (Color != "")
            {
                System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml(Color);
                _sheet.Cells[rowIndex, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                _sheet.Cells[rowIndex, col].Style.Fill.BackgroundColor.SetColor(colFromHex);

            }
        }
        protected int pintarcabeceras(List<string> _cabecera, ExcelWorksheet _Sheet, string header)
        {
            int finish = -1;
            foreach (var x in _cabecera)
            {
                finish++;
                string letra = convertNumberToLetter(finish);
                setHeader1(_Sheet, letra + "7:" + letra + "7", x, false);
            }
            _texto_sin_borde_Titulo1(_Sheet, convertNumberToLetter(0) + "1:" + convertNumberToLetter(finish) + "1", header, System.Drawing.Color.White, System.Drawing.Color.Black, "Calibri");
            return finish;
        }
        public string convertNumberToLetter(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var value = "";
            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];
            value += letters[index % letters.Length];
            return value;
        }
        public void setHeader1(ExcelWorksheet _sheet, string celda, string texto, Boolean mergue)
        {
            _sheet.Cells[celda].Value = texto;
            _sheet.Cells[celda].Style.Font.Name = "Calibri";
            _sheet.Cells[celda].Style.Font.Size = 12;
            _sheet.Cells[celda].Style.Font.Bold = true;
            _sheet.Cells[celda].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            _sheet.Cells[celda].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            _sheet.Cells[celda].Merge = mergue;
            _sheet.Cells[celda].Style.WrapText = false;
            _sheet.Cells[celda].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[celda].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[celda].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[celda].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[celda].Style.Fill.PatternType = ExcelFillStyle.Solid;
            System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#b4c6e7");
            _sheet.Cells[celda].Style.Fill.PatternType = ExcelFillStyle.Solid;
            _sheet.Cells[celda].Style.Fill.BackgroundColor.SetColor(colFromHex);
        }
        internal void _texto_sin_borde_Titulo1(ExcelWorksheet _sheet, String _range, String _text, System.Drawing.Color _Backcolor, System.Drawing.Color _fontColor, string fontName = "Calibri")
        {
            _sheet.Cells[_range].Value = _text;
            //_sheet.Cells[_range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //_sheet.Cells[_range].Style.Fill.BackgroundColor.SetColor(_Backcolor);
            //_sheet.Cells[_range].Style.Font.Color.SetColor(_fontColor);
            System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#72aea5");
            _sheet.Cells[_range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            _sheet.Cells[_range].Style.Fill.BackgroundColor.SetColor(colFromHex);
            _sheet.Cells[_range].Style.Font.Bold = true;
            _sheet.Cells[_range].Merge = true;
            _sheet.Cells[_range].Style.WrapText = false;
            //_sheet.Cells. = false;
            _sheet.Cells[_range].Style.Font.Size = 18;
            _sheet.Cells[_range].Style.Font.Name = fontName;
            _sheet.Cells[_range].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            _sheet.Cells[_range].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
    }
}
