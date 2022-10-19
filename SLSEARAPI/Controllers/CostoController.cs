using OfficeOpenXml;
using OfficeOpenXml.Style;
using SLSEARAPI.DataLayer;
using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace SLSEARAPI.Controllers
{
    public class CostoController : ApiController
    {
        CostoDL costoDL;
        public CostoController()
        {
            costoDL = new CostoDL();
        }

        [HttpPost]
        [ActionName("ListarCosto")]
        public List<Costo> ListarCosto(Costo costo)
        {
            try
            {
                return costoDL.ListarCosto(costo);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarActividad")]
        public List<Actividad> ListarActividad(Actividad actividad)
        {
            try
            {
                return costoDL.ListarActividad(actividad);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarCosto")]
        public Costo InsertarCosto(Costo costo)
        {
            try
            {
                return costoDL.InsertarCosto(costo);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ExportarCosto")]
        public HttpResponseMessage ExportarCosto(Cronograma cronograma)
        {
            try
            {
                String NombreReporte = "Costo";

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


                    List<string> _cabecera = new List<string>();

                    _cabecera.Add("Nro");
                    _cabecera.Add("Actividades");
                    _cabecera.Add("Descripcion");
                    _cabecera.Add("Unidad Medida");
                    _cabecera.Add("Cantidad");
                    _cabecera.Add("Costo Unitario");
                    _cabecera.Add("SubTotal");

                    int finish = pintarcabeceras(_cabecera, _genericSheet, "COSTO");

                    _genericSheet.Cells["A3:D3"].Value = "3.1 COSTO DE INVERSIÓN POR COMPONENTE / ACTIVIDAD / GASTO ELEGIBLE";
                    _genericSheet.Cells["A3:D3"].Merge = true;
                    _genericSheet.Cells["A3:D3"].Style.Locked = true;
                    System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#72aea5");
                    _genericSheet.Cells["A3:D3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["A3:D3"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    //pintar componente 1

                    //List<Componente> componentes = cronogramaDL.ListarComponentes(cronograma);
                    DataTable componentes = costoDL.ListarComponentesRpt(cronograma);
                    int rowIndexComp = 8;
                    int colcomp = 1;

                    if (componentes.Rows.Count > 0)
                    {
                        for (int i = 0; i < componentes.Rows.Count; i++)
                        {
                            //------------------------------------------------------------------------------------------------------------------------------------------------------
                            // Componentes
                            _texto_row(_genericSheet, rowIndexComp, colcomp++, (i + 1).ToString(), "#fff2cc");
                            _texto_row(_genericSheet, rowIndexComp, colcomp++, componentes.Rows[i][2], "#fff2cc");
                            _texto_row(_genericSheet, rowIndexComp, colcomp++, componentes.Rows[i][6], "#fff2cc");
                            _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#fff2cc");
                            _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#fff2cc");
                            _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#fff2cc");
                            _texto_row(_genericSheet, rowIndexComp, colcomp++, componentes.Rows[i][5], "#fff2cc");

                            for (int j = 10; j < componentes.Columns.Count; j++)
                            {
                                _texto_row_fecha(_genericSheet, 6, colcomp, Convert.ToDateTime(componentes.Columns[j].ColumnName), "#b4c6e7");
                                //_genericSheet.Rows[6].Height = 40;
                                _texto_row1(_genericSheet, 7, colcomp, j - 9, "#b4c6e7");
                                _texto_row1(_genericSheet, rowIndexComp, colcomp, componentes.Rows[i][j].ToString() == "" ? "-" : componentes.Rows[i][j].ToString(), "#fff2cc");
                                colcomp++;
                            }
                            rowIndexComp++;
                            colcomp = 1;
                            //------------------------------------------------------------------------------------------------------------------------------------------------------
                            // Actividades
                            Actividad actividad = new Actividad();
                            actividad.iCodIdentificacion = Convert.ToInt32(componentes.Rows[i][0].ToString());
                            actividad.iCodExtensionista = cronograma.iCodExtensionista;
                            DataTable listaactividades = costoDL.ListarActividadesPorComponente(actividad);
                            for (int k = 0; k < listaactividades.Rows.Count; k++)
                            {
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, (i + 1).ToString() + "." + (k + 1).ToString(), "#FFFFFF");
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, listaactividades.Rows[k][2], "#FFFFFF");
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, listaactividades.Rows[k][3], "#FFFFFF");
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, listaactividades.Rows[k][4], "#FFFFFF");
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, listaactividades.Rows[k][5], "#FFFFFF");
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#FFFFFF");
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, listaactividades.Rows[k][6], "#FFFFFF");
                                for (int z = 7; z < listaactividades.Columns.Count; z++)
                                {
                                    //_texto_row_fecha(_genericSheet, 6, colcomp, Convert.ToDateTime(componentes.Columns[j].ColumnName), "#b4c6e7");
                                    //_genericSheet.Rows[6].Height = 40;
                                    //_texto_row1(_genericSheet, 7, colcomp, j - 8, "#b4c6e7");
                                    if (listaactividades.Rows[k][z].ToString() == "")
                                    {
                                        _texto_row1(_genericSheet, rowIndexComp, colcomp, listaactividades.Rows[k][z].ToString(), "#E2EFDA");
                                    }
                                    else
                                    {
                                        _texto_row1(_genericSheet, rowIndexComp, colcomp, listaactividades.Rows[k][z].ToString(), "#00B050");
                                    }

                                    colcomp++;
                                }
                                rowIndexComp++;
                                colcomp = 1;
                                //------------------------------------------------------------------------------------------------------------------------------------------------------
                                //Costos Por Actividad
                                actividad = new Actividad();
                                actividad.iCodActividad = Convert.ToInt32(listaactividades.Rows[k][0].ToString());
                                actividad.iCodExtensionista = cronograma.iCodExtensionista;
                                DataTable listaCostos = costoDL.ListarCostosPorActividad(actividad);
                                for (int m = 0; m < listaCostos.Rows.Count; m++)
                                {
                                    _texto_row(_genericSheet, rowIndexComp, colcomp++, (i + 1).ToString() + "." + (k + 1).ToString() + "." + (m + 1).ToString(), "#FFFFFF");
                                    _texto_row(_genericSheet, rowIndexComp, colcomp++, listaCostos.Rows[m][1], "#FFFFFF");
                                    _texto_row(_genericSheet, rowIndexComp, colcomp++, listaCostos.Rows[m][2], "#FFFFFF");
                                    _texto_row(_genericSheet, rowIndexComp, colcomp++, listaCostos.Rows[m][3], "#FFFFFF");
                                    _texto_row(_genericSheet, rowIndexComp, colcomp++, listaCostos.Rows[m][4], "#FFFFFF");
                                    _texto_row(_genericSheet, rowIndexComp, colcomp++, listaCostos.Rows[m][5], "#FFFFFF");
                                    _texto_row(_genericSheet, rowIndexComp, colcomp++, listaCostos.Rows[m][6], "#FFFFFF");
                                    //_texto_row(_genericSheet, rowIndexComp, colcomp++, listaCostos.Rows[m][6], "#FFFFFF");
                                    for (int n = 7; n < listaCostos.Columns.Count; n++)
                                    {
                                        //_texto_row_fecha(_genericSheet, 6, colcomp, Convert.ToDateTime(componentes.Columns[j].ColumnName), "#b4c6e7");
                                        //_genericSheet.Rows[6].Height = 40;
                                        //_texto_row1(_genericSheet, 7, colcomp, j - 8, "#b4c6e7");
                                        if (listaCostos.Rows[m][n].ToString() == "")
                                        {
                                            _texto_row1(_genericSheet, rowIndexComp, colcomp, listaCostos.Rows[m][n].ToString(), "#E2EFDA");
                                        }
                                        else
                                        {
                                            _texto_row1(_genericSheet, rowIndexComp, colcomp, listaCostos.Rows[m][n].ToString(), "#00B050");
                                        }

                                        colcomp++;
                                    }
                                    rowIndexComp++;
                                    colcomp = 1;
                                }

                            }
                            
                        }


                        // Suma de Totales de Meta de Compromisos
                        int totalmetas = 0;

                        //foreach (Componente item in componentes)
                        //{
                        //    totalmetas = totalmetas + int.Parse(item.vMeta);
                        //}
                    }

                    rowIndexComp++;
                    rowIndexComp++;
                    int indCabecera = 0;

                    if (componentes.Rows.Count > 0)
                    {
                        //_cabecera = new List<string>();

                        //_cabecera.Add("Nro");
                        //_cabecera.Add("Gasto elegible");
                        //_cabecera.Add("Subtotal");
                        ////b4c6e7
                        //finish = pintarcabeceras(_cabecera, _genericSheet, "COSTO");

                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "Nro", "#b4c6e7");
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "Gasto elegible", "#b4c6e7");
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#b4c6e7");
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#b4c6e7");
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#b4c6e7");
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#b4c6e7");
                        _genericSheet.Cells["B" + rowIndexComp.ToString() + ":F" + rowIndexComp.ToString()].Merge = true;
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "Subtotal", "#b4c6e7");
                        rowIndexComp++;
                        colcomp = 1;
                        for (int i = 0; i < componentes.Rows.Count; i++)
                        {
                            
                            _texto_row(_genericSheet, rowIndexComp, colcomp++, (i + 1).ToString(), "#fff2cc");
                            _texto_row(_genericSheet, rowIndexComp, colcomp++, componentes.Rows[i][6], "#fff2cc");
                            _genericSheet.Cells["B" + rowIndexComp.ToString() + ":F" + rowIndexComp.ToString()].Merge = true;
                            colcomp = 7;
                            _texto_row(_genericSheet, rowIndexComp, colcomp++, componentes.Rows[i][5], "#fff2cc");


                            for (int j = 10; j < componentes.Columns.Count; j++)
                            {
                                if (indCabecera == 0)
                                {
                                    _texto_row_fecha(_genericSheet, rowIndexComp - 1, colcomp, Convert.ToDateTime(componentes.Columns[j].ColumnName), "#b4c6e7");
                                }
                                _texto_row1(_genericSheet, rowIndexComp, colcomp, componentes.Rows[i][j].ToString() == "" ? "-" : componentes.Rows[i][j].ToString(), "#fff2cc");
                                colcomp++;
                            }
                            indCabecera = 1;
                            rowIndexComp++;
                            colcomp = 1;

                            //------------------------------------------------------------------------------------------------------------------------------------------------------
                            // Gastos elegibles

                            for (int k = 1; k < 4; k++)
                            {
                                Actividad actividad = new Actividad();
                                actividad.iCodIdentificacion = Convert.ToInt32(componentes.Rows[i][0].ToString());
                                actividad.iopcion = k;
                                actividad.iCodExtensionista = cronograma.iCodExtensionista;
                                DataTable listaGastosElegibles = costoDL.ListarCostosResumenPorComponente(actividad);
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, componentes.Rows[i][6], "#FFFFFF");
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, listaGastosElegibles.Rows[0][0], "#FFFFFF");
                                _genericSheet.Cells["B" + rowIndexComp.ToString() + ":F" + rowIndexComp.ToString()].Merge = true;
                                colcomp = 7;
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, listaGastosElegibles.Rows[0][1], "#FFFFFF");

                                for (int z = 2; z < listaGastosElegibles.Columns.Count; z++)
                                {
                                    //_texto_row_fecha(_genericSheet, 6, colcomp, Convert.ToDateTime(componentes.Columns[j].ColumnName), "#b4c6e7");
                                    //_genericSheet.Rows[6].Height = 40;
                                    //_texto_row1(_genericSheet, 7, colcomp, j - 8, "#b4c6e7");
                                    if (listaGastosElegibles.Rows[0][z].ToString() == "" || listaGastosElegibles.Rows[0][z].ToString() == "0")
                                    {
                                        _texto_row1(_genericSheet, rowIndexComp, colcomp, "-", "#E2EFDA");
                                    }
                                    else
                                    {
                                        _texto_row1(_genericSheet, rowIndexComp, colcomp, listaGastosElegibles.Rows[0][z].ToString(), "#00B050");
                                    }

                                    colcomp++;
                                }
                                rowIndexComp++;
                                colcomp = 1;

                            }
                          


                        }
                    }

                    rowIndexComp++;
                    rowIndexComp++;
                    indCabecera = 0;

                    if (componentes.Rows.Count > 0)
                    {
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "Nro", "#b4c6e7");
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "Gasto elegible", "#b4c6e7");
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#b4c6e7");
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#b4c6e7");
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#b4c6e7");
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "", "#b4c6e7");
                        _genericSheet.Cells["B" + rowIndexComp.ToString() + ":F" + rowIndexComp.ToString()].Merge = true;
                        _texto_row(_genericSheet, rowIndexComp, colcomp++, "Subtotal", "#b4c6e7");
                        rowIndexComp++;
                        //colcomp = 1;
                        for (int i = 0; i < 1; i++)
                        {

                            
                            for (int j = 10; j < componentes.Columns.Count; j++)
                            {
                                if (indCabecera == 0)
                                {
                                    _texto_row_fecha(_genericSheet, rowIndexComp - 1 , colcomp, Convert.ToDateTime(componentes.Columns[j].ColumnName), "#b4c6e7");
                                }
                                colcomp++;
                            }
                            indCabecera = 1;
                            colcomp = 1;

                            //------------------------------------------------------------------------------------------------------------------------------------------------------
                            // Gastos elegibles

                            for (int k = 1; k < 4; k++)
                            {
                                Actividad actividad = new Actividad();
                                actividad.iopcion = k;
                                actividad.iCodExtensionista = cronograma.iCodExtensionista;
                                DataTable listaGastosElegibles = costoDL.ListarCostosResumenGeneral(actividad);
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, k, "#FFFFFF");
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, listaGastosElegibles.Rows[0][0], "#FFFFFF");
                                _genericSheet.Cells["B" + rowIndexComp.ToString() + ":F" + rowIndexComp.ToString()].Merge = true;
                                colcomp = 7;
                                _texto_row(_genericSheet, rowIndexComp, colcomp++, listaGastosElegibles.Rows[0][1], "#FFFFFF");

                                for (int z = 2; z < listaGastosElegibles.Columns.Count; z++)
                                {
                                    if (listaGastosElegibles.Rows[0][z].ToString() == "" || listaGastosElegibles.Rows[0][z].ToString() == "0")
                                    {
                                        _texto_row1(_genericSheet, rowIndexComp, colcomp, "-", "#E2EFDA");
                                    }
                                    else
                                    {
                                        _texto_row1(_genericSheet, rowIndexComp, colcomp, listaGastosElegibles.Rows[0][z].ToString(), "#00B050");
                                    }

                                    colcomp++;
                                }
                                rowIndexComp++;
                                colcomp = 1;

                            }



                        }
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

        public string convertNumberToLetter(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var value = "";
            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];
            value += letters[index % letters.Length];
            return value;
        }

        internal void _texto_row_fecha(ExcelWorksheet _sheet, int rowIndex, int col, object _text, string formato = "", string fontName = "Calibri")
        {
            _sheet.Cells[rowIndex, col].Value = _text;
            _sheet.Cells[rowIndex, col].Merge = true;
            _sheet.Columns[col].Width = 3;
            _sheet.Cells[rowIndex, col].Style.WrapText = false;
            _sheet.Cells[rowIndex, col].Style.Font.Name = fontName;
            _sheet.Cells[rowIndex, col].Style.Font.Size = 12;
            _sheet.Cells[rowIndex, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            _sheet.Cells[rowIndex, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[rowIndex, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[rowIndex, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#b4c6e7");
            _sheet.Cells[rowIndex, col].Style.Fill.PatternType = ExcelFillStyle.Solid;

            _sheet.Cells[rowIndex, col].Style.Fill.BackgroundColor.SetColor(colFromHex);
            if (formato != "")
            {
                _sheet.Cells[rowIndex, col].Style.Numberformat.Format = "d-mmm";
                //_sheet.Cells[rowIndex, col].Style.TextRotation = 90;

            }
        }

        internal void _texto_row1(ExcelWorksheet _sheet, int rowIndex, int col, object _text, string Color = "", string fontName = "Calibri")
        {
            _sheet.Cells[rowIndex, col].Value = _text;
            _sheet.Cells[rowIndex, col].Merge = true;
            _sheet.Columns[col].Width = 3;
            _sheet.Cells[rowIndex, col].Style.Font.Name = fontName;
            _sheet.Cells[rowIndex, col].Style.Font.Size = 12;
            _sheet.Cells[rowIndex, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            _sheet.Cells[rowIndex, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[rowIndex, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[rowIndex, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[rowIndex, col].Style.WrapText = false;
            //_sheet.Columns[col].AutoFit();

            if (Color != "")
            {
                System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml(Color);
                _sheet.Cells[rowIndex, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                _sheet.Cells[rowIndex, col].Style.Fill.BackgroundColor.SetColor(colFromHex);

            }
        }
    }
}