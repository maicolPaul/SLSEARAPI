using iTextSharp.text;
using Microsoft.Ajax.Utilities;
using OfficeOpenXml;
using OfficeOpenXml.Style;
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
using System.Web.UI.WebControls;

namespace SLSEARAPI.Controllers
{
    public class HitoController : ApiController
    {
        HitosDL hitosDL;

        private Exception ex;
        public HitoController()
        {

            hitosDL = new HitosDL();
        }

        [HttpPost]
        [ActionName("InsertarHito")]
        public Hito InsertarHito(Hito entidad)
        {
            try
            {
                return hitosDL.InsertarHito(entidad);
            }
            catch (Exception)
            {
                throw;
            }

        }
        [HttpPost]
        [ActionName("InsertarProductorEje")]
        public PorductorEjecucionTecnica InsertarProductorEje(PorductorEjecucionTecnica porductorEjecucionTecnica)
        {
            try
            {
                return hitosDL.InsertarProductorEje(porductorEjecucionTecnica);
            }
            catch (Exception)
            {
                throw;
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


        [HttpPost]
        [ActionName("ExportarEjecucionTecFin")]
        public HttpResponseMessage ExportarEjecucionTecFin(FichaTecnica fichaTecnica)
        {
            try
            {
                int indColumnaResumen1 = 0;

                String NombreReporte = "Ejecucion_Tecnica_Financiera";

                using (var excelPackage = new ExcelPackage())
                {
                    #region ConfigPestaña
                    excelPackage.Workbook.Properties.Author = NombreReporte;
                    excelPackage.Workbook.Properties.Title = NombreReporte;
                    var _genericSheet = excelPackage.Workbook.Worksheets.Add(NombreReporte);
                    _genericSheet.View.ShowGridLines = false;
                    _genericSheet.View.ZoomScale = 100;
                    _genericSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                    _genericSheet.PrinterSettings.FitToPage = true;
                    _genericSheet.PrinterSettings.Orientation = eOrientation.Landscape;
                    _genericSheet.View.PageBreakView = true;
                    #endregion

                    #region CABECERA

                    _genericSheet.Cells["A1:AZ1"].Value = "";
                    _genericSheet.Cells["A1:AZ1"].Merge = true;
                    _genericSheet.Cells["A1:AZ1"].Style.Locked = true;
                    _genericSheet.Cells["A1:AZ1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                    _genericSheet.Cells["A1:AZ1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["A1:AZ1"].Style.Font.Size = 18;
                    _genericSheet.Cells["A1:AZ1"].Style.Font.Bold = true;
                    _genericSheet.Cells["A1:AZ1"].Style.Font.Name = "Calibri";
                    _genericSheet.Rows[4].Height = 21;
                    _genericSheet.Cells["A1:AZ1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    _genericSheet.Cells["A2:F2"].Value = "EJECUCIÓN TÉCNICA - FINANCIERA";
                    _genericSheet.Cells["A2:F2"].Merge = true;
                    _genericSheet.Cells["A2:F2"].Style.Locked = true;
                    _genericSheet.Cells["A2:F2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                    _genericSheet.Cells["A2:F2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["A2:F2"].Style.Font.Size = 18;
                    _genericSheet.Cells["A2:F2"].Style.Font.Bold = true;
                    _genericSheet.Cells["A2:F2"].Style.Font.Name = "Calibri";
                    _genericSheet.Rows[4].Height = 21;
                    _genericSheet.Cells["A2:F2"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;


                    _genericSheet.Cells["G2:AZ2"].Value = "";
                    _genericSheet.Cells["G2:AZ2"].Merge = true;
                    _genericSheet.Cells["G2:AZ2"].Style.Locked = true;
                    _genericSheet.Cells["G2:AZ2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                    _genericSheet.Cells["G2:AZ2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["G2:AZ2"].Style.Font.Size = 18;
                    _genericSheet.Cells["G2:AZ2"].Style.Font.Bold = true;
                    _genericSheet.Cells["G2:AZ2"].Style.Font.Name = "Calibri";
                    _genericSheet.Rows[4].Height = 21;
                    _genericSheet.Cells["G2:AZ2"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    _genericSheet.Cells["A3:AZ3"].Value = "";
                    _genericSheet.Cells["A3:AZ3"].Merge = true;
                    _genericSheet.Cells["A3:AZ3"].Style.Locked = true;
                    _genericSheet.Cells["A3:AZ3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                    _genericSheet.Cells["A3:AZ3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["A3:AZ3"].Style.Font.Size = 18;
                    _genericSheet.Cells["A3:AZ3"].Style.Font.Bold = true;
                    _genericSheet.Cells["A3:AZ3"].Style.Font.Name = "Calibri";
                    _genericSheet.Rows[4].Height = 21;
                    _genericSheet.Cells["A3:AZ3"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    _genericSheet.Cells["A4:F6"].Value = "";
                    _genericSheet.Cells["A4:F6"].Merge = true;
                    _genericSheet.Cells["A4:F6"].Style.Locked = true;
                    _genericSheet.Cells["A4:F6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                    _genericSheet.Cells["A4:F6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["A4:F6"].Style.Font.Size = 18;
                    _genericSheet.Cells["A4:F6"].Style.Font.Bold = true;
                    _genericSheet.Cells["A4:F6"].Style.Font.Name = "Calibri";
                    _genericSheet.Rows[4].Height = 21;
                    _genericSheet.Cells["A4:F6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    #endregion

                    #region
                    DataTable cortesDT = hitosDL.ListarCortes(fichaTecnica);
                    int indFila = 4;
                    int indColumna = 7;
                    int indColumnaFin = 0;
                    if (cortesDT.Rows.Count > 0)
                    {
                         
                        for (int i = 0; i < cortesDT.Rows.Count; i++)
                        {
                            indColumnaFin = indColumna + 5; 
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Value = (i + 1).ToString() + "° Entregable";
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Merge = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Locked = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Size = 11;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Bold = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Name = "Calibri";
                            _genericSheet.Rows[4].Height = 14.4;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            _genericSheet.Column(1).Style.Locked = true;
                            _genericSheet.Workbook.Protection.LockWindows = true;
                            _genericSheet.Workbook.Protection.LockStructure = true;
                            indColumna += 6;

                            indColumnaResumen1 = indColumna;

                        }

                        indFila++;
                        indColumna = 7;
                        indColumnaFin = 0;
                        for (int i = 0; i < cortesDT.Rows.Count; i++)
                        {
                            indColumnaFin = indColumna + 5;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Value = cortesDT.Rows[i][1].ToString();
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Merge = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Locked = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Size = 11;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Bold = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Name = "Calibri";
                            _genericSheet.Rows[4].Height = 14.4;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            _genericSheet.Column(1).Style.Locked = true;
                            _genericSheet.Workbook.Protection.LockWindows = true;
                            _genericSheet.Workbook.Protection.LockStructure = true;
                            indColumna += 6;
                        }

                        indFila++;
                        indColumna = 7;
                        indColumnaFin = 0;
                        for (int i = 0; i < cortesDT.Rows.Count; i++)
                        {
                            indColumnaFin = indColumna + 2;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Value = "Tecnico";
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Merge = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Locked = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFF00");
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Size = 11;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Bold = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Name = "Calibri";
                            _genericSheet.Rows[4].Height = 14.4;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            _genericSheet.Column(1).Style.Locked = true;
                            _genericSheet.Workbook.Protection.LockWindows = true;
                            _genericSheet.Workbook.Protection.LockStructure = true;
                            indColumna += 3;
                            
                            indColumnaFin = indColumna + 2;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Value = "Financiero (S/.)";
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Merge = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Locked = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B050");
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Size = 11;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Bold = true;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Name = "Calibri";
                            _genericSheet.Rows[4].Height = 14.4;
                            _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            _genericSheet.Column(1).Style.Locked = true;
                            _genericSheet.Workbook.Protection.LockWindows = true;
                            _genericSheet.Workbook.Protection.LockStructure = true;
                            indColumna += 3;
                        }
                    }

                    #region CabeceraComponente
                    _genericSheet.Cells["B7:B7"].Value = "N°";
                    _genericSheet.Cells["B7:B7"].Merge = true;
                    _genericSheet.Cells["B7:B7"].Style.Locked = true;
                    _genericSheet.Cells["B7:B7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                    _genericSheet.Cells["B7:B7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["B7:B7"].Style.Font.Size = 11;
                    _genericSheet.Cells["B7:B7"].Style.Font.Bold = true;
                    _genericSheet.Cells["B7:B7"].Style.Font.Name = "Calibri";
                    _genericSheet.Columns[2].Width = 5.33;
                    _genericSheet.Cells["B7:B7"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Cells["B7:B7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    _genericSheet.Cells["C7:C7"].Value = "Gasto Elegible";
                    _genericSheet.Cells["C7:C7"].Merge = true;
                    _genericSheet.Cells["C7:C7"].Style.Locked = true;
                    _genericSheet.Cells["C7:C7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                    _genericSheet.Cells["C7:C7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["C7:C7"].Style.Font.Size = 11;
                    _genericSheet.Cells["C7:C7"].Style.Font.Bold = true;
                    _genericSheet.Cells["C7:C7"].Style.Font.Name = "Calibri";
                    _genericSheet.Columns[3].Width = 13.33;
                    _genericSheet.Cells["C7:C7"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Cells["C7:C7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    _genericSheet.Cells["D7:D7"].Value = "Descripción";
                    _genericSheet.Cells["D7:D7"].Merge = true;
                    _genericSheet.Cells["D7:D7"].Style.Locked = true;
                    _genericSheet.Cells["D7:D7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                    _genericSheet.Cells["D7:D7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["D7:D7"].Style.Font.Size = 11;
                    _genericSheet.Cells["D7:D7"].Style.Font.Bold = true;
                    _genericSheet.Cells["D7:D7"].Style.Font.Name = "Calibri";
                    _genericSheet.Columns[4].Width = 32.67;
                    _genericSheet.Cells["D7:D7"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Cells["D7:D7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    _genericSheet.Cells["E7:E7"].Value = "Unidad de Medida";
                    _genericSheet.Cells["E7:E7"].Merge = true;
                    _genericSheet.Cells["E7:E7"].Style.Locked = true;
                    _genericSheet.Cells["E7:E7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                    _genericSheet.Cells["E7:E7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["E7:E7"].Style.Font.Size = 11;
                    _genericSheet.Cells["E7:E7"].Style.Font.Bold = true;
                    _genericSheet.Cells["E7:E7"].Style.Font.Name = "Calibri";
                    _genericSheet.Columns[5].Width = 10.78;
                    _genericSheet.Cells["E7:E7"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Cells["E7:E7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    _genericSheet.Cells["E7:E7"].Style.WrapText = true;
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    _genericSheet.Cells["F7:F7"].Value = "Meta";
                    _genericSheet.Cells["F7:F7"].Merge = true;
                    _genericSheet.Cells["F7:F7"].Style.Locked = true;
                    _genericSheet.Cells["F7:F7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                    _genericSheet.Cells["F7:F7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["F7:F7"].Style.Font.Size = 11;
                    _genericSheet.Cells["F7:F7"].Style.Font.Bold = true;
                    _genericSheet.Cells["F7:F7"].Style.Font.Name = "Calibri";
                    _genericSheet.Columns[5].Width = 10.78;
                    _genericSheet.Cells["F7:F7"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Cells["F7:F7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;


                    indFila++;
                    indColumna = 7;
                    indColumnaFin = 0;
                    for (int i = 0; i < cortesDT.Rows.Count; i++)
                    {
                        indColumnaFin = indColumna;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Value = "Prog";
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Merge = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Locked = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Size = 11;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Bold = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Name = "Calibri";
                        //_genericSheet.Rows[4].Height = 14.4;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _genericSheet.Column(1).Style.Locked = true;
                        _genericSheet.Workbook.Protection.LockWindows = true;
                        _genericSheet.Workbook.Protection.LockStructure = true;
                        indColumna += 1;

                        indColumnaFin = indColumna;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Value = "Ejec";
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Merge = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Locked = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Size = 11;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Bold = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Name = "Calibri";
                        //_genericSheet.Rows[4].Height = 14.4;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _genericSheet.Column(1).Style.Locked = true;
                        _genericSheet.Workbook.Protection.LockWindows = true;
                        _genericSheet.Workbook.Protection.LockStructure = true;
                        indColumna += 1;

                        indColumnaFin = indColumna;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Value = "%";
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Merge = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Locked = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Size = 11;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Bold = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Name = "Calibri";
                        //_genericSheet.Rows[4].Height = 14.4;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _genericSheet.Column(1).Style.Locked = true;
                        _genericSheet.Workbook.Protection.LockWindows = true;
                        _genericSheet.Workbook.Protection.LockStructure = true;
                        indColumna += 1;

                        indColumnaFin = indColumna;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Value = "Prog";
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Merge = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Locked = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Size = 11;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Bold = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Name = "Calibri";
                       // _genericSheet.Rows[4].Height = 14.4;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _genericSheet.Column(1).Style.Locked = true;
                        _genericSheet.Workbook.Protection.LockWindows = true;
                        _genericSheet.Workbook.Protection.LockStructure = true;
                        indColumna += 1;

                        indColumnaFin = indColumna;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Value = "Ejec";
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Merge = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Locked = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Size = 11;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Bold = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Name = "Calibri";
                        //_genericSheet.Rows[4].Height = 14.4;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _genericSheet.Column(1).Style.Locked = true;
                        _genericSheet.Workbook.Protection.LockWindows = true;
                        _genericSheet.Workbook.Protection.LockStructure = true;
                        indColumna += 1;

                        indColumnaFin = indColumna;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Value = "%";
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Merge = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Locked = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#9BC2E6");
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Size = 11;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Bold = true;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Font.Name = "Calibri";
                        //_genericSheet.Rows[4].Height = 14.4;
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        _genericSheet.Cells[indFila, indColumna, indFila, indColumnaFin].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _genericSheet.Column(1).Style.Locked = true;
                        _genericSheet.Workbook.Protection.LockWindows = true;
                        _genericSheet.Workbook.Protection.LockStructure = true;
                        indColumna += 1;
                    }
                    _genericSheet.Rows[7].Height = 35;
                    _genericSheet.Columns[1].Width = 1;

                    #endregion

                    #region Componentes
                    DataSet componentesDS = hitosDL.ListarComponentes(fichaTecnica);
                    DataTable componentesDT = componentesDS.Tables[0];
                    if (componentesDT.Rows.Count > 0)
                    {
                        //indColumna = 2;
                        string indComponente = "";
                        int correlativo = 0;
                        for (int i = 0; i < componentesDT.Rows.Count; i++)
                        {
                            if (componentesDT.Rows[i][13].ToString() != indComponente)
                            {
                                #region Componente
                                correlativo++;
                                indComponente = componentesDT.Rows[i][13].ToString();
                                indFila++;
                                indColumna = 2;
                                //indColumnaFin = indColumna + 5;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = correlativo.ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][0].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][1].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][2].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][3].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][5].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][6].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][7].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][8].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][9].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][10].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                #endregion

                                
                            }
                            else {
                                //indFila++;
                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][5].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][6].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][7].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][8].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][9].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;

                                indColumna++;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = componentesDT.Rows[i][10].ToString();
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = true;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                //_genericSheet.Rows[4].Height = 14.4;
                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                _genericSheet.Column(1).Style.Locked = true;
                                _genericSheet.Workbook.Protection.LockWindows = true;
                                _genericSheet.Workbook.Protection.LockStructure = true;
                            }

                            if (componentesDT.Rows.Count > (i + 1))
                            {
                                if (componentesDT.Rows[i + 1][13].ToString() != indComponente)
                                {
                                    #region PintarActividades
                                    FichaTecnica ft = new FichaTecnica();
                                    ft.iCodExtensionista = Convert.ToInt32(indComponente);
                                    DataSet actividadesDS = hitosDL.ListarActividades(ft);
                                    DataTable actividadesDT = actividadesDS.Tables[0];
                                    if (actividadesDT.Rows.Count > 0)
                                    {
                                        int correlativoActividad = 0;
                                        //indComponente = componentesDT.Rows[i][11].ToString();

                                        string indActividad = "";
                                        for (int j = 0; j < actividadesDT.Rows.Count; j++)
                                        {
                                            if (actividadesDT.Rows[j][0].ToString() != indActividad)
                                            {
                                                indActividad = actividadesDT.Rows[j][0].ToString();
                                                correlativoActividad++;
                                                indFila++;
                                                indColumna = 2;
                                                
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = correlativo.ToString() + "." + correlativoActividad.ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][0].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][1].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][2].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][3].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][5].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][6].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][7].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][8].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][9].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][10].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;
                                            }
                                            else
                                            {
                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][5].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][6].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][7].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][8].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][9].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][10].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;
                                            }

                                        }
                                    }
                                    #endregion
                                }
                            }
                            else if (componentesDT.Rows.Count == (i + 1))
                            {
                                //if (componentesDT.Rows[i + 1][13].ToString() != indComponente)
                                //{
                                    #region PintarActividades
                                    FichaTecnica ft = new FichaTecnica();
                                    ft.iCodExtensionista = Convert.ToInt32(indComponente);
                                    DataSet actividadesDS = hitosDL.ListarActividades(ft);
                                    DataTable actividadesDT = actividadesDS.Tables[0];
                                    if (actividadesDT.Rows.Count > 0)
                                    {
                                        int correlativoActividad = 0;
                                        //indComponente = componentesDT.Rows[i][11].ToString();

                                        string indActividad = "";
                                        for (int j = 0; j < actividadesDT.Rows.Count; j++)
                                        {
                                            if (actividadesDT.Rows[j][0].ToString() != indActividad)
                                            {
                                                indActividad = actividadesDT.Rows[j][0].ToString();
                                                correlativoActividad++;
                                                indFila++;
                                                indColumna = 2;
                                                // #D0CECE
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = correlativo.ToString() + "." + correlativoActividad.ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][0].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][1].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][2].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][3].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][5].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][6].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][7].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][8].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][9].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][10].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;
                                            }
                                            else
                                            {
                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][5].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][6].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][7].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][8].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][9].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;

                                                indColumna++;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][10].ToString();
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                                //_genericSheet.Rows[4].Height = 14.4;
                                                _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                                _genericSheet.Column(1).Style.Locked = true;
                                                _genericSheet.Workbook.Protection.LockWindows = true;
                                                _genericSheet.Workbook.Protection.LockStructure = true;
                                            }

                                        }
                                    }
                                    #endregion
                                //}
                            }
                            
                        }


                    }

                    DataTable componentesResumenDT = componentesDS.Tables[2];

                    #region CabeceraResumen
                    indFila = 5;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Value = "";
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Merge = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Locked = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B0F0");
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Font.Size = 11;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Font.Bold = false;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Font.Name = "Calibri";
                    //_genericSheet.Rows[4].Height = 14.4;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    /**************************************************************************************************************************************************/

                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Value = "";
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Merge = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Locked = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B0F0");
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Font.Size = 11;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Font.Bold = false;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Font.Name = "Calibri";
                    //_genericSheet.Rows[4].Height = 14.4;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    /**************************************************************************************************************************************************/
                    /**************************************************************************************************************************************************/
                    /**************************************************************************************************************************************************/

                    indFila++;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Value = "Acumulado Tecnico";
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Merge = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Locked = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B0F0");
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Font.Size = 11;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Font.Bold = false;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Font.Name = "Calibri";
                    //_genericSheet.Rows[4].Height = 14.4;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 4].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    /**************************************************************************************************************************************************/

                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Value = "Acumulado Financiero";
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Merge = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Locked = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B0F0");
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Font.Size = 11;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Font.Bold = false;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Font.Name = "Calibri";
                    //_genericSheet.Rows[4].Height = 14.4;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 8].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    /**************************************************************************************************************************************************/
                    /**************************************************************************************************************************************************/
                    /**************************************************************************************************************************************************/

                    indFila++;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Value = "Prog";
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Merge = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Locked = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B0F0");
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Font.Size = 11;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Font.Bold = false;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Font.Name = "Calibri";
                    //_genericSheet.Rows[4].Height = 14.4;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Value = "Ejec";
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Merge = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Locked = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B0F0");
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Font.Size = 11;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Font.Bold = false;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Font.Name = "Calibri";
                    //_genericSheet.Rows[4].Height = 14.4;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Value = "% al Total";
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Merge = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Locked = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B0F0");
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Font.Size = 11;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Font.Bold = false;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Font.Name = "Calibri";
                    //_genericSheet.Rows[4].Height = 14.4;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    /**************************************************************************************************************************************************/

                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Value = "Prog";
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Merge = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Locked = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B0F0");
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Font.Size = 11;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Font.Bold = false;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Font.Name = "Calibri";
                    //_genericSheet.Rows[4].Height = 14.4;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Value = "Ejec";
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Merge = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Locked = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B0F0");
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Font.Size = 11;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Font.Bold = false;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Font.Name = "Calibri";
                    //_genericSheet.Rows[4].Height = 14.4;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Value = "% al Total";
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Merge = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Locked = true;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    colFromHex = System.Drawing.ColorTranslator.FromHtml("#00B0F0");
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Font.Size = 11;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Font.Bold = false;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Font.Name = "Calibri";
                    //_genericSheet.Rows[4].Height = 14.4;
                    _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;
                    #endregion

                    if (componentesResumenDT.Rows.Count > 0)
                    {
                        int indComponente = 0;
                        for (int i = 0; i < componentesResumenDT.Rows.Count; i++)
                        {
                            #region ComponentesResumen
                            indComponente = Convert.ToInt32(componentesResumenDT.Rows[i][10].ToString());
                            indFila++;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Value = componentesResumenDT.Rows[i][4].ToString();
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Merge = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Locked = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Font.Size = 11;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Font.Bold = false;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Font.Name = "Calibri";
                            //_genericSheet.Rows[4].Height = 14.4;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 2, indFila, indColumnaResumen1 + 2].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            _genericSheet.Column(1).Style.Locked = true;
                            _genericSheet.Workbook.Protection.LockWindows = true;
                            _genericSheet.Workbook.Protection.LockStructure = true;

                            _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Value = componentesResumenDT.Rows[i][5].ToString();
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Merge = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Locked = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Font.Size = 11;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Font.Bold = false;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Font.Name = "Calibri";
                            //_genericSheet.Rows[4].Height = 14.4;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 3, indFila, indColumnaResumen1 + 3].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            _genericSheet.Column(1).Style.Locked = true;
                            _genericSheet.Workbook.Protection.LockWindows = true;
                            _genericSheet.Workbook.Protection.LockStructure = true;

                            _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Value = componentesResumenDT.Rows[i][6].ToString();
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Merge = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Locked = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Font.Size = 11;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Font.Bold = false;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Font.Name = "Calibri";
                            //_genericSheet.Rows[4].Height = 14.4;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 4, indFila, indColumnaResumen1 + 4].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            _genericSheet.Column(1).Style.Locked = true;
                            _genericSheet.Workbook.Protection.LockWindows = true;
                            _genericSheet.Workbook.Protection.LockStructure = true;

                            /**************************************************************************************************************************************************/

                            _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Value = componentesResumenDT.Rows[i][7].ToString();
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Merge = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Locked = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Font.Size = 11;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Font.Bold = false;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Font.Name = "Calibri";
                            //_genericSheet.Rows[4].Height = 14.4;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 6, indFila, indColumnaResumen1 + 6].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            _genericSheet.Column(1).Style.Locked = true;
                            _genericSheet.Workbook.Protection.LockWindows = true;
                            _genericSheet.Workbook.Protection.LockStructure = true;

                            _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Value = componentesResumenDT.Rows[i][8].ToString();
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Merge = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Locked = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Font.Size = 11;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Font.Bold = false;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Font.Name = "Calibri";
                            //_genericSheet.Rows[4].Height = 14.4;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 7, indFila, indColumnaResumen1 + 7].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            _genericSheet.Column(1).Style.Locked = true;
                            _genericSheet.Workbook.Protection.LockWindows = true;
                            _genericSheet.Workbook.Protection.LockStructure = true;

                            _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Value = componentesResumenDT.Rows[i][9].ToString();
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Merge = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Locked = true;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFE699");
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Font.Size = 11;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Font.Bold = false;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Font.Name = "Calibri";
                            //_genericSheet.Rows[4].Height = 14.4;
                            _genericSheet.Cells[indFila, indColumnaResumen1 + 8, indFila, indColumnaResumen1 + 8].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            _genericSheet.Column(1).Style.Locked = true;
                            _genericSheet.Workbook.Protection.LockWindows = true;
                            _genericSheet.Workbook.Protection.LockStructure = true;
                            #endregion

                            if (componentesResumenDT.Rows.Count > (i + 1))
                            {
                                #region PintarActividades
                                    FichaTecnica ft = new FichaTecnica();
                                    ft.iCodExtensionista = Convert.ToInt32(indComponente);
                                    DataSet actividadesDS = hitosDL.ListarActividades(ft);
                                    DataTable actividadesDT = actividadesDS.Tables[1];
                                    if (actividadesDT.Rows.Count > 0)
                                    {
                                        int correlativoActividad = 0;
                                        //indComponente = componentesDT.Rows[i][11].ToString();

                                        string indActividad = "";
                                        for (int j = 0; j < actividadesDT.Rows.Count; j++)
                                        {
                                            indActividad = actividadesDT.Rows[j][0].ToString();
                                            correlativoActividad++;
                                            indFila++;
                                            indColumna = indColumnaResumen1 + 2;

                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][4].ToString();
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                            //_genericSheet.Rows[4].Height = 14.4;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                            _genericSheet.Column(1).Style.Locked = true;
                                            _genericSheet.Workbook.Protection.LockWindows = true;
                                            _genericSheet.Workbook.Protection.LockStructure = true;

                                            indColumna++;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][5].ToString();
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                            //_genericSheet.Rows[4].Height = 14.4;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                            _genericSheet.Column(1).Style.Locked = true;
                                            _genericSheet.Workbook.Protection.LockWindows = true;
                                            _genericSheet.Workbook.Protection.LockStructure = true;

                                            indColumna++;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][6].ToString();
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                            //_genericSheet.Rows[4].Height = 14.4;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                            _genericSheet.Column(1).Style.Locked = true;
                                            _genericSheet.Workbook.Protection.LockWindows = true;
                                            _genericSheet.Workbook.Protection.LockStructure = true;

                                            indColumna++;
                                            indColumna++;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][7].ToString();
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                            //_genericSheet.Rows[4].Height = 14.4;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                            _genericSheet.Column(1).Style.Locked = true;
                                            _genericSheet.Workbook.Protection.LockWindows = true;
                                            _genericSheet.Workbook.Protection.LockStructure = true;

                                            indColumna++;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][8].ToString();
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                            //_genericSheet.Rows[4].Height = 14.4;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                            _genericSheet.Column(1).Style.Locked = true;
                                            _genericSheet.Workbook.Protection.LockWindows = true;
                                            _genericSheet.Workbook.Protection.LockStructure = true;

                                            indColumna++;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][9].ToString();
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                            //_genericSheet.Rows[4].Height = 14.4;
                                            _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                            _genericSheet.Column(1).Style.Locked = true;
                                            _genericSheet.Workbook.Protection.LockWindows = true;
                                            _genericSheet.Workbook.Protection.LockStructure = true;

                                            

                                        }
                                    }
                                    #endregion
                            }
                            else if (componentesResumenDT.Rows.Count == (i + 1))
                            {
                                #region PintarActividades
                                FichaTecnica ft = new FichaTecnica();
                                ft.iCodExtensionista = Convert.ToInt32(indComponente);
                                DataSet actividadesDS = hitosDL.ListarActividades(ft);
                                DataTable actividadesDT = actividadesDS.Tables[1];
                                if (actividadesDT.Rows.Count > 0)
                                {
                                    int correlativoActividad = 0;
                                    //indComponente = componentesDT.Rows[i][11].ToString();

                                    string indActividad = "";
                                    for (int j = 0; j < actividadesDT.Rows.Count; j++)
                                    {
                                        indActividad = actividadesDT.Rows[j][0].ToString();
                                        correlativoActividad++;
                                        indFila++;
                                        indColumna = indColumnaResumen1 + 2;

                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][4].ToString();
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                        //_genericSheet.Rows[4].Height = 14.4;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                        _genericSheet.Column(1).Style.Locked = true;
                                        _genericSheet.Workbook.Protection.LockWindows = true;
                                        _genericSheet.Workbook.Protection.LockStructure = true;

                                        indColumna++;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][5].ToString();
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                        //_genericSheet.Rows[4].Height = 14.4;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                        _genericSheet.Column(1).Style.Locked = true;
                                        _genericSheet.Workbook.Protection.LockWindows = true;
                                        _genericSheet.Workbook.Protection.LockStructure = true;

                                        indColumna++;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][6].ToString();
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                        //_genericSheet.Rows[4].Height = 14.4;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                        _genericSheet.Column(1).Style.Locked = true;
                                        _genericSheet.Workbook.Protection.LockWindows = true;
                                        _genericSheet.Workbook.Protection.LockStructure = true;

                                        indColumna++;
                                        indColumna++;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][7].ToString();
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                        //_genericSheet.Rows[4].Height = 14.4;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                        _genericSheet.Column(1).Style.Locked = true;
                                        _genericSheet.Workbook.Protection.LockWindows = true;
                                        _genericSheet.Workbook.Protection.LockStructure = true;

                                        indColumna++;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][8].ToString();
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                        //_genericSheet.Rows[4].Height = 14.4;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                        _genericSheet.Column(1).Style.Locked = true;
                                        _genericSheet.Workbook.Protection.LockWindows = true;
                                        _genericSheet.Workbook.Protection.LockStructure = true;

                                        indColumna++;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Value = actividadesDT.Rows[j][9].ToString();
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Merge = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Locked = true;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#D0CECE");
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Size = 11;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Bold = false;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Font.Name = "Calibri";
                                        //_genericSheet.Rows[4].Height = 14.4;
                                        _genericSheet.Cells[indFila, indColumna, indFila, indColumna].Style.Fill.BackgroundColor.SetColor(colFromHex);
                                        _genericSheet.Column(1).Style.Locked = true;
                                        _genericSheet.Workbook.Protection.LockWindows = true;
                                        _genericSheet.Workbook.Protection.LockStructure = true;

                                    }
                                }
                                #endregion
                            }
                        }
                    }
                    #endregion

                    #endregion

                    /******************************************************************************************/

                    var x = excelPackage.GetAsByteArray();
                    NombreReporte = "Ejecucion_Tecnica_Financiera";
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

        protected int pintarcabeceras_(List<string> _cabecera, ExcelWorksheet _Sheet, string header, int fila, int columnas)
        {
            int finish = -1;
            foreach (var x in _cabecera)
            {
                finish++;
                string letra = convertNumberToLetter(finish);
                setHeader1(_Sheet, letra + fila.ToString() + ":" + letra + fila.ToString(), x, false);
            }
            _texto_sin_borde_Titulo1(_Sheet, convertNumberToLetter(0) + "1:" + convertNumberToLetter(finish) + "1", header, System.Drawing.Color.White, System.Drawing.Color.Black, "Calibri");
            return finish;
        }
    }
}