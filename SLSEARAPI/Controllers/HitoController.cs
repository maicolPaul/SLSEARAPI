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
                    DataTable componentesDT = hitosDL.ListarComponentes(fichaTecnica);
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
                                    DataTable actividadesDT = hitosDL.ListarActividades(ft);
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
                                    DataTable actividadesDT = hitosDL.ListarActividades(ft);
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
                    #endregion

                    #endregion

                    /***********************************************************************************/
                    #region
                    //#region InformacionGeneral

                    ///*****************************************************************************************/

                    //#region CABECERA

                    //_genericSheet.Cells["A4:J4"].Value = "SERVICIOS DE EXTENSIÓN AGRARIA RURAL";
                    //_genericSheet.Cells["A4:J4"].Merge = true;
                    //_genericSheet.Cells["A4:J4"].Style.Locked = true;
                    //_genericSheet.Cells["A4:J4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A4:J4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Cells["A4:J4"].Style.Font.Size = 16;
                    //_genericSheet.Cells["A4:J4"].Style.Font.Bold = true;
                    //_genericSheet.Cells["A4:J4"].Style.Font.Name = "Univers Light";
                    //_genericSheet.Rows[4].Height = 21;
                    //_genericSheet.Cells["A4:J4"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A6:J6"].Value = "FICHA TÉCNICA";
                    //_genericSheet.Cells["A6:J6"].Merge = true;
                    //_genericSheet.Cells["A6:J6"].Style.Locked = true;
                    //_genericSheet.Cells["A6:J6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#548235");
                    //_genericSheet.Cells["A6:J6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Cells["A6:J6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //System.Drawing.Color colFromHex1 = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A6:J6"].Style.Font.Size = 22;
                    //_genericSheet.Cells["A6:J6"].Style.Font.Bold = true;
                    //_genericSheet.Cells["A6:J6"].Style.Font.Name = "Univers";
                    //_genericSheet.Cells["A6:J6"].Style.Font.Color.SetColor(colFromHex1);
                    //_genericSheet.Rows[6].Height = 28.2;
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A8:J8"].Value = "INFORMACIÓN GENERAL";
                    //_genericSheet.Cells["A8:J8"].Merge = true;
                    //_genericSheet.Cells["A8:J8"].Style.Locked = true;
                    //_genericSheet.Cells["A8:J8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A8:J8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Cells["A8:J8"].Style.Font.Size = 18;
                    //_genericSheet.Cells["A8:J8"].Style.Font.Bold = true;
                    //_genericSheet.Cells["A8:J8"].Style.Font.Name = "Univers Condensed Light";
                    //_genericSheet.Cells["A8:J8"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Rows[8].Height = 22.8;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //#endregion

                    ///*****************************************************************************************/

                    //#region 1.1 DATOS GENERALES DEL SEAR

                    //_genericSheet.Cells["A10:H10"].Value = "1.1 DATOS GENERALES DEL SEAR";
                    //_genericSheet.Cells["A10:H10"].Merge = true;
                    //_genericSheet.Cells["A10:H10"].Style.Locked = true;
                    //_genericSheet.Cells["A10:H10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#548235");
                    //_genericSheet.Cells["A10:H10"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Cells["A10:H10"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A12:D12"].Value = "1.1.1 Nombre del SEAR";
                    //_genericSheet.Cells["A12:D12"].Merge = true;
                    //_genericSheet.Cells["A12:D12"].Style.Locked = true;
                    //_genericSheet.Cells["A12:D12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A12:D12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A12:D12"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Cells["A12:D12"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E12:H12"].Value = ficha[0].vNombreSearT1;
                    //_genericSheet.Cells["E12:H12"].Merge = true;
                    //_genericSheet.Cells["E12:H12"].Style.Locked = true;
                    //_genericSheet.Cells["E12:H12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E12:H12"].Style.WrapText = true;
                    //_genericSheet.Cells["E12:H12"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Rows[12].Height = 87.6;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E12:H12"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A14:D14"].Value = "1.1.2 Naturaleza de la Intervención";
                    //_genericSheet.Cells["A14:D14"].Merge = true;
                    //_genericSheet.Cells["A14:D14"].Style.Locked = true;
                    //_genericSheet.Cells["A14:D14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A14:D14"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A14:D14"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E14:H14"].Value = ficha[0].vNaturalezaIntervencionT1;
                    //_genericSheet.Cells["E14:H14"].Merge = true;
                    //_genericSheet.Cells["E14:H14"].Style.Locked = true;
                    //_genericSheet.Cells["E14:H14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E14:H14"].Style.Fill.PatternType = ExcelFillStyle.Solid;

                    ////_genericSheet.Rows[12].Height = 87.6;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E14:H14"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A16:D16"].Value = "1.1.3 Sub Sector";
                    //_genericSheet.Cells["A16:D16"].Merge = true;
                    //_genericSheet.Cells["A16:D16"].Style.Locked = true;
                    //_genericSheet.Cells["A16:D16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A16:D16"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A16:D16"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E16:E16"].Value = ficha[0].vSubSectorT1;
                    //_genericSheet.Cells["E16:E16"].Merge = true;
                    //_genericSheet.Cells["E16:E16"].Style.Locked = true;
                    //_genericSheet.Cells["E16:E16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E16:E16"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[12].Height = 87.6;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E16:E16"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A18:D18"].Value = "1.1.4 Nombre de la cadena productiva priorizada";
                    //_genericSheet.Cells["A18:D18"].Merge = true;
                    //_genericSheet.Cells["A18:D18"].Style.Locked = true;
                    //_genericSheet.Cells["A18:D18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A18:D18"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A18:D18"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F18:H18"].Value = ficha[0].vCadenaProductivaT1;
                    //_genericSheet.Cells["F18:H18"].Merge = true;
                    //_genericSheet.Cells["F18:H18"].Style.Locked = true;
                    //_genericSheet.Cells["F18:H18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F18:H18"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Rows[18].Height = 27.6;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F18:H18"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A20:D20"].Value = "1.1.5 Proceso de la cadena productiva";
                    //_genericSheet.Cells["A20:D20"].Merge = true;
                    //_genericSheet.Cells["A20:D20"].Style.Locked = true;
                    //_genericSheet.Cells["A20:D20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A20:D20"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A20:D20"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E20:E20"].Value = ficha[0].vProcesoProductivaT1;
                    //_genericSheet.Cells["E20:E20"].Merge = true;
                    //_genericSheet.Cells["E20:E20"].Style.Locked = true;
                    //_genericSheet.Cells["E20:E20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E20:E20"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[12].Height = 87.6;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E20:E20"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A22:D22"].Value = "1.1.6 Línea prioritaria";
                    //_genericSheet.Cells["A22:D22"].Merge = true;
                    //_genericSheet.Cells["A22:D22"].Style.Locked = true;
                    //_genericSheet.Cells["A22:D22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A22:D22"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A22:D22"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F22:G22"].Value = ficha[0].vLineaPrioritariaT1;
                    //_genericSheet.Cells["F22:G22"].Merge = true;
                    //_genericSheet.Cells["F22:G22"].Style.Locked = true;
                    //_genericSheet.Cells["F22:G22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F22:G22"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[12].Height = 87.6;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F22:G22"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A24:D24"].Value = "1.1.7 Producto o servicio a ampliar / mejorar / recuperar";
                    //_genericSheet.Cells["A24:D24"].Merge = true;
                    //_genericSheet.Cells["A24:D24"].Style.Locked = true;
                    //_genericSheet.Cells["A24:D24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A24:D24"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A24:D24"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F24:H24"].Value = ficha[0].vProductoServicioAmpliarT1;
                    //_genericSheet.Cells["F24:H24"].Merge = true;
                    //_genericSheet.Cells["F24:H24"].Style.Locked = true;
                    //_genericSheet.Cells["F24:H24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F24:H24"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F24:H24"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A26:D26"].Value = "1.1.8 Localización del servicio";
                    //_genericSheet.Cells["A26:D26"].Merge = true;
                    //_genericSheet.Cells["A26:D26"].Style.Locked = true;
                    //_genericSheet.Cells["A26:D26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A26:D26"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A26:D26"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E26:E26"].Value = "Región";
                    //_genericSheet.Cells["E26:E26"].Merge = true;
                    //_genericSheet.Cells["E26:E26"].Style.Locked = true;
                    //_genericSheet.Cells["E26:E26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E26:E26"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E26:E26"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F26:F26"].Value = ficha[0].vNomDepartamento;
                    //_genericSheet.Cells["F26:F26"].Merge = true;
                    //_genericSheet.Cells["F26:F26"].Style.Locked = true;
                    //_genericSheet.Cells["F26:F26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F26:F26"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F26:F26"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["E27:E27"].Value = "Provincia";
                    //_genericSheet.Cells["E27:E27"].Merge = true;
                    //_genericSheet.Cells["E27:E27"].Style.Locked = true;
                    //_genericSheet.Cells["E27:E27"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E27:E27"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E27:E27"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F27:F27"].Value = ficha[0].vNomProvincia;
                    //_genericSheet.Cells["F27:F27"].Merge = true;
                    //_genericSheet.Cells["F27:F27"].Style.Locked = true;
                    //_genericSheet.Cells["F27:F27"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F27:F27"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F27:F27"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["E28:E28"].Value = "Distrito";
                    //_genericSheet.Cells["E28:E28"].Merge = true;
                    //_genericSheet.Cells["E28:E28"].Style.Locked = true;
                    //_genericSheet.Cells["E28:E28"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E28:E28"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E28:E28"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F28:F28"].Value = ficha[0].vNomDistrito;
                    //_genericSheet.Cells["F28:F28"].Merge = true;
                    //_genericSheet.Cells["F28:F28"].Style.Locked = true;
                    //_genericSheet.Cells["F28:F28"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F28:F28"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F28:F28"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["E29:E29"].Value = "Localidad";
                    //_genericSheet.Cells["E29:E29"].Merge = true;
                    //_genericSheet.Cells["E29:E29"].Style.Locked = true;
                    //_genericSheet.Cells["E29:E29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E29:E29"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E29:E29"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F29:H29"].Value = ficha[0].vLocalidadT1;
                    //_genericSheet.Cells["F29:H29"].Merge = true;
                    //_genericSheet.Cells["F29:H29"].Style.Locked = true;
                    //_genericSheet.Cells["F29:H29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F29:H29"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F29:H29"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A31:D31"].Value = "1.1.9 Ubicación del servicio";
                    //_genericSheet.Cells["A31:D31"].Merge = true;
                    //_genericSheet.Cells["A31:D31"].Style.Locked = true;
                    //_genericSheet.Cells["A31:D31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A31:D31"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A31:D31"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E31:E31"].Value = "Zona UTM";
                    //_genericSheet.Cells["E31:E31"].Merge = true;
                    //_genericSheet.Cells["E31:E31"].Style.Locked = true;
                    //_genericSheet.Cells["E31:E31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E31:E31"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E31:E31"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F31:H31"].Value = ficha[0].vZonaUTMT1;
                    //_genericSheet.Cells["F31:H31"].Merge = true;
                    //_genericSheet.Cells["F31:H31"].Style.Locked = true;
                    //_genericSheet.Cells["F31:H31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F31:H31"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F31:H31"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["E32:E32"].Value = "Coordenadas UTM (Norte)";
                    //_genericSheet.Cells["E32:E32"].Merge = true;
                    //_genericSheet.Cells["E32:E32"].Style.Locked = true;
                    //_genericSheet.Cells["E32:E32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E32:E32"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E32:E32"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F32:H32"].Value = ficha[0].vCoordenadasUTMNorteT1;
                    //_genericSheet.Cells["F32:H32"].Merge = true;
                    //_genericSheet.Cells["F32:H32"].Style.Locked = true;
                    //_genericSheet.Cells["F32:H32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F32:H32"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F32:H32"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["E33:E33"].Value = "Coordenadas UTM (Este)";
                    //_genericSheet.Cells["E33:E33"].Merge = true;
                    //_genericSheet.Cells["E33:E33"].Style.Locked = true;
                    //_genericSheet.Cells["E33:E33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E33:E33"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E33:E33"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F33:H33"].Value = ficha[0].vCoordenadasUTMEsteT1;
                    //_genericSheet.Cells["F33:H33"].Merge = true;
                    //_genericSheet.Cells["F33:H33"].Style.Locked = true;
                    //_genericSheet.Cells["F33:H33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F33:H33"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F33:H33"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A35:D35"].Value = "1.1.10 Fecha de inicio y fin del servicio";
                    //_genericSheet.Cells["A35:D35"].Merge = true;
                    //_genericSheet.Cells["A35:D35"].Style.Locked = true;
                    //_genericSheet.Cells["A35:D35"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A35:D35"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A35:D35"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E35:E35"].Value = "Inicio";
                    //_genericSheet.Cells["E35:E35"].Merge = true;
                    //_genericSheet.Cells["E35:E35"].Style.Locked = true;
                    //_genericSheet.Cells["E35:E35"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E35:E35"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E35:E35"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F35:F35"].Value = ficha[0].dFechaInicioServicioT1;
                    //_genericSheet.Cells["F35:F35"].Merge = true;
                    //_genericSheet.Cells["F35:F35"].Style.Locked = true;
                    //_genericSheet.Cells["F35:F35"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F35:F35"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F35:F35"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["E36:E36"].Value = "Fin";
                    //_genericSheet.Cells["E36:E36"].Merge = true;
                    //_genericSheet.Cells["E36:E36"].Style.Locked = true;
                    //_genericSheet.Cells["E36:E36"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E36:E36"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E36:E36"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F36:F36"].Value = ficha[0].dFechaFinServicioT1;
                    //_genericSheet.Cells["F36:F36"].Merge = true;
                    //_genericSheet.Cells["F36:F36"].Style.Locked = true;
                    //_genericSheet.Cells["F36:F36"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F36:F36"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F36:F36"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A38:D38"].Value = "1.1.11 Duración (días)";
                    //_genericSheet.Cells["A38:D38"].Merge = true;
                    //_genericSheet.Cells["A38:D38"].Style.Locked = true;
                    //_genericSheet.Cells["A38:D38"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A38:D38"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A38:D38"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E38:E38"].Value = ficha[0].TotalDias;
                    //_genericSheet.Cells["E38:E38"].Merge = true;
                    //_genericSheet.Cells["E38:E38"].Style.Locked = true;
                    //_genericSheet.Cells["E38:E38"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["E38:E38"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E38:E38"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //#endregion

                    ///*****************************************************************************************/

                    //#region 1.2 DATOS DE LA ENTIDAD PROMOTORA

                    //_genericSheet.Cells["A40:H40"].Value = "1.2 DATOS DE LA ENTIDAD PROMOTORA";
                    //_genericSheet.Cells["A40:H40"].Merge = true;
                    //_genericSheet.Cells["A40:H40"].Style.Locked = true;
                    //_genericSheet.Cells["A40:H40"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Cells["A40:H40"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#548235");
                    //_genericSheet.Cells["A40:H40"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A42:D42"].Value = "1.2.1 Nombre de la entidad proponente (Agencia Agraria)";
                    //_genericSheet.Cells["A42:D42"].Merge = true;
                    //_genericSheet.Cells["A42:D42"].Style.Locked = true;
                    //_genericSheet.Cells["A42:D42"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A42:D42"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A42:D42"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F42:H42"].Value = "1.2.5 Correo electrónico";
                    //_genericSheet.Cells["F42:H42"].Merge = true;
                    //_genericSheet.Cells["F42:H42"].Style.Locked = true;
                    //_genericSheet.Cells["F42:H42"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F42:H42"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["F42:H42"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["A43:D43"].Value = ficha[0].vNombreEntidadProponenteT2;
                    //_genericSheet.Cells["A43:D43"].Merge = true;
                    //_genericSheet.Cells["A43:D43"].Style.Locked = true;
                    //_genericSheet.Cells["A43:D43"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A43:D43"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A43:D43"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F43:H43"].Value = ficha[0].vCorreoElectronicoT2;
                    //_genericSheet.Cells["F43:H43"].Merge = true;
                    //_genericSheet.Cells["F43:H43"].Style.Locked = true;
                    //_genericSheet.Cells["F43:H43"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F43:H43"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F43:H43"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A44:D44"].Value = "1.2.2 Nombre de la Dirección Regional a la que pertenece";
                    //_genericSheet.Cells["A44:D44"].Merge = true;
                    //_genericSheet.Cells["A44:D44"].Style.Locked = true;
                    //_genericSheet.Cells["A44:D44"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A44:D44"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A44:D44"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F44:H44"].Value = "1.2.6 Nombre del director de la Agencia Agraria";
                    //_genericSheet.Cells["F44:H44"].Merge = true;
                    //_genericSheet.Cells["F44:H44"].Style.Locked = true;
                    //_genericSheet.Cells["F44:H44"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F44:H44"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["F44:H44"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["A45:D45"].Value = ficha[0].vNombreDireccionPerteneceT2;
                    //_genericSheet.Cells["A45:D45"].Merge = true;
                    //_genericSheet.Cells["A45:D45"].Style.Locked = true;
                    //_genericSheet.Cells["A45:D45"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A45:D45"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A45:D45"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F45:H45"].Value = ficha[0].vNombreDirectorAgenciaAgrariaT2;
                    //_genericSheet.Cells["F45:H45"].Merge = true;
                    //_genericSheet.Cells["F45:H45"].Style.Locked = true;
                    //_genericSheet.Cells["F45:H45"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F45:H45"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F45:H45"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A46:D46"].Value = "1.2.3 Dirección";
                    //_genericSheet.Cells["A46:D46"].Merge = true;
                    //_genericSheet.Cells["A46:D46"].Style.Locked = true;
                    //_genericSheet.Cells["A46:D46"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A46:D46"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A46:D46"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F46:H46"].Value = "1.2.7 Dirección Zonal AGRORURAL";
                    //_genericSheet.Cells["F46:H46"].Merge = true;
                    //_genericSheet.Cells["F46:H46"].Style.Locked = true;
                    //_genericSheet.Cells["F46:H46"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F46:H46"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["F46:H46"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["A47:D47"].Value = ficha[0].vDireccionT2;
                    //_genericSheet.Cells["A47:D47"].Merge = true;
                    //_genericSheet.Cells["A47:D47"].Style.Locked = true;
                    //_genericSheet.Cells["A47:D47"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A47:D47"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A47:D47"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F47:H47"].Value = ficha[0].vDireccionZonaAgroruralT2;
                    //_genericSheet.Cells["F47:H47"].Merge = true;
                    //_genericSheet.Cells["F47:H47"].Style.Locked = true;
                    //_genericSheet.Cells["F47:H47"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F47:H47"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F47:H47"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A48:B48"].Value = "1.2.4 Teléfono";
                    //_genericSheet.Cells["A48:B48"].Merge = true;
                    //_genericSheet.Cells["A48:B48"].Style.Locked = true;
                    //_genericSheet.Cells["A48:B48"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A48:B48"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A48:B48"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["A49:B49"].Value = ficha[0].vTelefonoT2;
                    //_genericSheet.Cells["A49:B49"].Merge = true;
                    //_genericSheet.Cells["A49:B49"].Style.Locked = true;
                    //_genericSheet.Cells["A49:B49"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A49:B49"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ////_genericSheet.Rows[22].Height = 44;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A49:B49"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //#endregion

                    ///*****************************************************************************************/

                    //#region 1.3 DATOS DEL PROVEEDOR		

                    //_genericSheet.Cells["A51:H51"].Value = "1.3 DATOS DEL PROVEEDOR";
                    //_genericSheet.Cells["A51:H51"].Merge = true;
                    //_genericSheet.Cells["A51:H51"].Style.Locked = true;
                    //_genericSheet.Cells["A51:H51"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Cells["A51:H51"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#548235");
                    //_genericSheet.Cells["A51:H51"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A53:D53"].Value = "1.3.1 Tipo de personeria juridica";
                    //_genericSheet.Cells["A53:D53"].Merge = true;
                    //_genericSheet.Cells["A53:D53"].Style.Locked = true;
                    //_genericSheet.Cells["A53:D53"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A53:D53"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A53:D53"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F53:F53"].Value = "1.3.7 Teléfono";
                    //_genericSheet.Cells["F53:F53"].Merge = true;
                    //_genericSheet.Cells["F53:F53"].Style.Locked = true;
                    //_genericSheet.Cells["F53:F53"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F53:F53"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["F53:F53"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["H53:H53"].Value = "1.3.8 Celular";
                    //_genericSheet.Cells["H53:H53"].Merge = true;
                    //_genericSheet.Cells["H53:H53"].Style.Locked = true;
                    //_genericSheet.Cells["H53:H53"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H53:H53"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["H53:H53"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A54:C54"].Value = ficha[0].TipoPersoneriaT3;
                    //_genericSheet.Cells["A54:C54"].Merge = true;
                    //_genericSheet.Cells["A54:C54"].Style.Locked = true;
                    //_genericSheet.Cells["A54:C54"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A54:C54"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A54:C54"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["F54:F54"].Value = ficha[0].vTelefonoT3;
                    //_genericSheet.Cells["F54:F54"].Merge = true;
                    //_genericSheet.Cells["F54:F54"].Style.Locked = true;
                    //_genericSheet.Cells["F54:F54"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F54:F54"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F54:F54"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["H54:H54"].Value = ficha[0].vCelularT3;
                    //_genericSheet.Cells["H54:H54"].Merge = true;
                    //_genericSheet.Cells["H54:H54"].Style.Locked = true;
                    //_genericSheet.Cells["H54:H54"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H54:H54"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["H54:H54"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A55:D55"].Value = "1.3.2 Nombre o razón social del proveedor";
                    //_genericSheet.Cells["A55:D55"].Merge = true;
                    //_genericSheet.Cells["A55:D55"].Style.Locked = true;
                    //_genericSheet.Cells["A55:D55"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A55:D55"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A55:D55"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F55:H55"].Value = "1.3.9 Correo electrónico";
                    //_genericSheet.Cells["F55:H55"].Merge = true;
                    //_genericSheet.Cells["F55:H55"].Style.Locked = true;
                    //_genericSheet.Cells["F55:H55"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F55:H55"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["F55:H55"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A56:D56"].Value = ficha[0].vNombreRazonSocialProveedorT3;
                    //_genericSheet.Cells["A56:D56"].Merge = true;
                    //_genericSheet.Cells["A56:D56"].Style.Locked = true;
                    //_genericSheet.Cells["A56:D56"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A56:D56"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A56:D56"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["F56:H56"].Value = ficha[0].vCorreoElectronicoT3;
                    //_genericSheet.Cells["F56:H56"].Merge = true;
                    //_genericSheet.Cells["F56:H56"].Style.Locked = true;
                    //_genericSheet.Cells["F56:H56"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F56:H56"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F56:H56"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A57:D57"].Value = "1.3.3 Nombre del proveedor o representante legal";
                    //_genericSheet.Cells["A57:D57"].Merge = true;
                    //_genericSheet.Cells["A57:D57"].Style.Locked = true;
                    //_genericSheet.Cells["A57:D57"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A57:D57"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A57:D57"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F57:H57"].Value = "1.3.10 Página web";
                    //_genericSheet.Cells["F57:H57"].Merge = true;
                    //_genericSheet.Cells["F57:H57"].Style.Locked = true;
                    //_genericSheet.Cells["F57:H57"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F57:H57"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["F57:H57"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A58:D58"].Value = ficha[0].vNombreRepresentanteLegalT3;
                    //_genericSheet.Cells["A58:D58"].Merge = true;
                    //_genericSheet.Cells["A58:D58"].Style.Locked = true;
                    //_genericSheet.Cells["A58:D58"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A58:D58"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A58:D58"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["F58:H58"].Value = ficha[0].vPaginaWebT3;
                    //_genericSheet.Cells["F58:H58"].Merge = true;
                    //_genericSheet.Cells["F58:H58"].Style.Locked = true;
                    //_genericSheet.Cells["F58:H58"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F58:H58"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F58:H58"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A59:B59"].Value = "1.3.4 N° DNI";
                    //_genericSheet.Cells["A59:B59"].Merge = true;
                    //_genericSheet.Cells["A59:B59"].Style.Locked = true;
                    //_genericSheet.Cells["A59:B59"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A59:B59"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A59:B59"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["D59:D59"].Value = "1.3.5 N° RUC";
                    //_genericSheet.Cells["D59:D59"].Merge = true;
                    //_genericSheet.Cells["D59:D59"].Style.Locked = true;
                    //_genericSheet.Cells["D59:D59"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["D59:D59"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["D59:D59"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F59:H59"].Value = "1.3.12 Especialidad o profesión del proveedor";
                    //_genericSheet.Cells["F59:H59"].Merge = true;
                    //_genericSheet.Cells["F59:H59"].Style.Locked = true;
                    //_genericSheet.Cells["F59:H59"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F59:H59"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["F59:H59"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A60:B60"].Value = ficha[0].vDniT3;
                    //_genericSheet.Cells["A60:B60"].Merge = true;
                    //_genericSheet.Cells["A60:B60"].Style.Locked = true;
                    //_genericSheet.Cells["A60:B60"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A60:B60"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A60:B60"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["D60:D60"].Value = ficha[0].vRucT3;
                    //_genericSheet.Cells["D60:D60"].Merge = true;
                    //_genericSheet.Cells["D60:D60"].Style.Locked = true;
                    //_genericSheet.Cells["D60:D60"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["D60:D60"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["D60:D60"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["F60:H60"].Value = ficha[0].vEpecialidadProveedorT3;
                    //_genericSheet.Cells["F60:H60"].Merge = true;
                    //_genericSheet.Cells["F60:H60"].Style.Locked = true;
                    //_genericSheet.Cells["F60:H60"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F60:H60"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F60:H60"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A61:C61"].Value = "1.3.6 Dirección";
                    //_genericSheet.Cells["A61:C61"].Merge = true;
                    //_genericSheet.Cells["A61:C61"].Style.Locked = true;
                    //_genericSheet.Cells["A61:C61"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A61:C61"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A61:C61"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F61:H61"].Value = "1.3.13 Tipo de proveedor";
                    //_genericSheet.Cells["F61:H61"].Merge = true;
                    //_genericSheet.Cells["F61:H61"].Style.Locked = true;
                    //_genericSheet.Cells["F61:H61"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F61:H61"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["F61:H61"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A62:D62"].Value = ficha[0].vDireccionT3;
                    //_genericSheet.Cells["A62:D62"].Merge = true;
                    //_genericSheet.Cells["A62:D62"].Style.Locked = true;
                    //_genericSheet.Cells["A62:D62"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A62:D62"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A62:D62"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["F60:H60"].Value = ficha[0].vProveedor;
                    //_genericSheet.Cells["F60:H60"].Merge = true;
                    //_genericSheet.Cells["F60:H60"].Style.Locked = true;
                    //_genericSheet.Cells["F60:H60"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F60:H60"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F60:H60"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //#endregion


                    ///*****************************************************************************************/

                    //#region 1.4 DATOS DE LOS BENEFICIARIOS	

                    //List<Productor> Organizacion = costoDL.SP_Listar_OrganizacionesRpt(fichaTecnica);

                    //_genericSheet.Cells["A64:H64"].Value = "1.4 DATOS DE LOS BENEFICIARIOS";
                    //_genericSheet.Cells["A64:H64"].Merge = true;
                    //_genericSheet.Cells["A64:H64"].Style.Locked = true;
                    //_genericSheet.Cells["A64:H64"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Cells["A64:H64"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#548235");
                    //_genericSheet.Cells["A64:H64"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A65:D65"].Value = "Organización 1";
                    //_genericSheet.Cells["A65:D65"].Merge = true;
                    //_genericSheet.Cells["A65:D65"].Style.Locked = true;
                    //_genericSheet.Cells["A65:D65"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A65:D65"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A65:D65"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E65:G65"].Value = "Organización 2";
                    //_genericSheet.Cells["E65:G65"].Merge = true;
                    //_genericSheet.Cells["E65:G65"].Style.Locked = true;
                    //_genericSheet.Cells["E65:G65"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E65:G65"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E65:G65"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["H65:J65"].Value = "Organización 3";
                    //_genericSheet.Cells["H65:J65"].Merge = true;
                    //_genericSheet.Cells["H65:J65"].Style.Locked = true;
                    //_genericSheet.Cells["H65:J65"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H65:J65"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["H65:J65"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A66:D66"].Value = "1.4.1 Nombre o razón social";
                    //_genericSheet.Cells["A66:D66"].Merge = true;
                    //_genericSheet.Cells["A66:D66"].Style.Locked = true;
                    //_genericSheet.Cells["A66:D66"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A66:D66"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A66:D66"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["E66:G66"].Value = "1.4.1 Nombre o razón social";
                    //_genericSheet.Cells["E66:G66"].Merge = true;
                    //_genericSheet.Cells["E66:G66"].Style.Locked = true;
                    //_genericSheet.Cells["E66:G66"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E66:G66"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E66:G66"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["H66:J66"].Value = "1.4.1 Nombre o razón social";
                    //_genericSheet.Cells["H66:J66"].Merge = true;
                    //_genericSheet.Cells["H66:J66"].Style.Locked = true;
                    //_genericSheet.Cells["H66:J66"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["H66:J66"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["H66:J66"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A67:D67"].Value = Organizacion[0].vNombreOrganizacion;
                    //_genericSheet.Cells["A67:D67"].Merge = true;
                    //_genericSheet.Cells["A67:D67"].Style.Locked = true;
                    //_genericSheet.Cells["A67:D67"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A67:D67"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A67:D67"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["E67:G67"].Value = Organizacion[1].vNombreOrganizacion;
                    //_genericSheet.Cells["E67:G67"].Merge = true;
                    //_genericSheet.Cells["E67:G67"].Style.Locked = true;
                    //_genericSheet.Cells["E67:G67"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["E67:G67"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E67:G67"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["H67:J67"].Value = Organizacion[2].vNombreOrganizacion;
                    //_genericSheet.Cells["H67:J67"].Merge = true;
                    //_genericSheet.Cells["H67:J67"].Style.Locked = true;
                    //_genericSheet.Cells["H67:J67"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H67:J67"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["H67:J67"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A68:D68"].Value = "1.4.2 Nombre del representante";
                    //_genericSheet.Cells["A68:D68"].Merge = true;
                    //_genericSheet.Cells["A68:D68"].Style.Locked = true;
                    //_genericSheet.Cells["A68:D68"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A68:D68"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A68:D68"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E68:G68"].Value = "1.4.2 Nombre del representante";
                    //_genericSheet.Cells["E68:G68"].Merge = true;
                    //_genericSheet.Cells["E68:G68"].Style.Locked = true;
                    //_genericSheet.Cells["E68:G68"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E68:G68"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E68:G68"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["H68:J68"].Value = "1.4.2 Nombre del representante";
                    //_genericSheet.Cells["H68:J68"].Merge = true;
                    //_genericSheet.Cells["H68:J68"].Style.Locked = true;
                    //_genericSheet.Cells["H68:J68"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["H68:J68"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["H68:J68"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A69:D69"].Value = Organizacion[0].vNombreRepresentante;
                    //_genericSheet.Cells["A69:D69"].Merge = true;
                    //_genericSheet.Cells["A69:D69"].Style.Locked = true;
                    //_genericSheet.Cells["A69:D69"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A69:D69"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A69:D69"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["E69:G69"].Value = Organizacion[1].vNombreRepresentante;
                    //_genericSheet.Cells["E69:G69"].Merge = true;
                    //_genericSheet.Cells["E69:G69"].Style.Locked = true;
                    //_genericSheet.Cells["E69:G69"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E69:G69"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E69:G69"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["H69:J69"].Value = Organizacion[2].vNombreRepresentante;
                    //_genericSheet.Cells["H69:J69"].Merge = true;
                    //_genericSheet.Cells["H69:J69"].Style.Locked = true;
                    //_genericSheet.Cells["H69:J69"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["H69:J69"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["H69:J69"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A70:D70"].Value = "1.4.3 N° RUC";
                    //_genericSheet.Cells["A70:D70"].Merge = true;
                    //_genericSheet.Cells["A70:D70"].Style.Locked = true;
                    //_genericSheet.Cells["A70:D70"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A70:D70"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A70:D70"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E70:G70"].Value = "1.4.3 N° RUC";
                    //_genericSheet.Cells["E70:G70"].Merge = true;
                    //_genericSheet.Cells["E70:G70"].Style.Locked = true;
                    //_genericSheet.Cells["E70:G70"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E70:G70"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E70:G70"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["H70:J70"].Value = "1.4.3 N° RUC";
                    //_genericSheet.Cells["H70:J70"].Merge = true;
                    //_genericSheet.Cells["H70:J70"].Style.Locked = true;
                    //_genericSheet.Cells["H70:J70"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["H70:J70"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["H70:J70"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A71:D71"].Value = Organizacion[0].vRucOrg;
                    //_genericSheet.Cells["A71:D71"].Merge = true;
                    //_genericSheet.Cells["A71:D71"].Style.Locked = true;
                    //_genericSheet.Cells["A71:D71"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A71:D71"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A71:D71"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["E71:G71"].Value = Organizacion[1].vRucOrg;
                    //_genericSheet.Cells["E71:G71"].Merge = true;
                    //_genericSheet.Cells["E71:G71"].Style.Locked = true;
                    //_genericSheet.Cells["E71:G71"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["E71:G71"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E71:G71"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["H71:J71"].Value = Organizacion[2].vRucOrg;
                    //_genericSheet.Cells["H71:J71"].Merge = true;
                    //_genericSheet.Cells["H71:J71"].Style.Locked = true;
                    //_genericSheet.Cells["H71:J71"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H71:J71"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["H71:J71"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A72:B72"].Value = "1.4.4 Teléfono";
                    //_genericSheet.Cells["A72:B72"].Merge = true;
                    //_genericSheet.Cells["A72:B72"].Style.Locked = true;
                    //_genericSheet.Cells["A72:B72"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A72:B72"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A72:B72"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["C72:C72"].Value = "1.4.5 Celuar";
                    //_genericSheet.Cells["C72:C72"].Merge = true;
                    //_genericSheet.Cells["C72:C72"].Style.Locked = true;
                    //_genericSheet.Cells["C72:C72"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["C72:C72"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["C72:C72"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["E72:E72"].Value = "1.4.4 Teléfono";
                    //_genericSheet.Cells["E72:E72"].Merge = true;
                    //_genericSheet.Cells["E72:E72"].Style.Locked = true;
                    //_genericSheet.Cells["E72:E72"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E72:E72"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E72:E72"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F72:F72"].Value = "1.4.5 Celuar";
                    //_genericSheet.Cells["F72:F72"].Merge = true;
                    //_genericSheet.Cells["F72:F72"].Style.Locked = true;
                    //_genericSheet.Cells["F72:F72"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["F72:F72"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["F72:F72"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["H72:H72"].Value = "1.4.4 Teléfono";
                    //_genericSheet.Cells["H72:H72"].Merge = true;
                    //_genericSheet.Cells["H72:H72"].Style.Locked = true;
                    //_genericSheet.Cells["H72:H72"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["H72:H72"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["H72:H72"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["I72:I72"].Value = "1.4.5 Celular";
                    //_genericSheet.Cells["I72:I72"].Merge = true;
                    //_genericSheet.Cells["I72:I72"].Style.Locked = true;
                    //_genericSheet.Cells["I72:I72"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["I72:I72"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["I72:I72"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A73:A73"].Value = Organizacion[0].vTelefonoOrg;
                    //_genericSheet.Cells["A73:A73"].Merge = true;
                    //_genericSheet.Cells["A73:A73"].Style.Locked = true;
                    //_genericSheet.Cells["A73:A73"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A73:A73"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A73:A73"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["B73:B73"].Value = Organizacion[0].vCelularOrg;
                    //_genericSheet.Cells["B73:B73"].Merge = true;
                    //_genericSheet.Cells["B73:B73"].Style.Locked = true;
                    //_genericSheet.Cells["B73:B73"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["B73:B73"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["B73:B73"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["E73:E73"].Value = Organizacion[1].vTelefonoOrg;
                    //_genericSheet.Cells["E73:E73"].Merge = true;
                    //_genericSheet.Cells["E73:E73"].Style.Locked = true;
                    //_genericSheet.Cells["E73:E73"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["E73:E73"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E73:E73"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["F73:F73"].Value = Organizacion[1].vCelularOrg;
                    //_genericSheet.Cells["F73:F73"].Merge = true;
                    //_genericSheet.Cells["F73:F73"].Style.Locked = true;
                    //_genericSheet.Cells["F73:F73"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F73:F73"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["F73:F73"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["H73:H73"].Value = Organizacion[2].vTelefonoOrg;
                    //_genericSheet.Cells["H73:H73"].Merge = true;
                    //_genericSheet.Cells["H73:H73"].Style.Locked = true;
                    //_genericSheet.Cells["H73:H73"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H73:H73"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["H73:H73"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["I73:I73"].Value = Organizacion[2].vCelularOrg;
                    //_genericSheet.Cells["I73:I73"].Merge = true;
                    //_genericSheet.Cells["I73:I73"].Style.Locked = true;
                    //_genericSheet.Cells["I73:I73"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["I73:I73"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["I73:I73"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A74:D74"].Value = "1.4.6 Dirección";
                    //_genericSheet.Cells["A74:D74"].Merge = true;
                    //_genericSheet.Cells["A74:D74"].Style.Locked = true;
                    //_genericSheet.Cells["A74:D74"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A74:D74"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A74:D74"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E74:G74"].Value = "1.4.6 Dirección";
                    //_genericSheet.Cells["E74:G74"].Merge = true;
                    //_genericSheet.Cells["E74:G74"].Style.Locked = true;
                    //_genericSheet.Cells["E74:G74"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E74:G74"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E74:G74"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["H74:J74"].Value = "1.4.6 Dirección";
                    //_genericSheet.Cells["H74:J74"].Merge = true;
                    //_genericSheet.Cells["H74:J74"].Style.Locked = true;
                    //_genericSheet.Cells["H74:J74"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H74:J74"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["H74:J74"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A75:D75"].Value = Organizacion[0].vDireccionOrg;
                    //_genericSheet.Cells["A75:D75"].Merge = true;
                    //_genericSheet.Cells["A75:D75"].Style.Locked = true;
                    //_genericSheet.Cells["A75:D75"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A75:D75"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A75:D75"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["E75:G75"].Value = Organizacion[1].vDireccionOrg;
                    //_genericSheet.Cells["E75:G75"].Merge = true;
                    //_genericSheet.Cells["E75:G75"].Style.Locked = true;
                    //_genericSheet.Cells["E75:G75"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["E75:G75"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E75:G75"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["H75:J75"].Value = Organizacion[2].vDireccionOrg;
                    //_genericSheet.Cells["H75:J75"].Merge = true;
                    //_genericSheet.Cells["H75:J75"].Style.Locked = true;
                    //_genericSheet.Cells["H75:J75"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H75:J75"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["H75:J75"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A76:D76"].Value = "1.4.7 Correo electrónico";
                    //_genericSheet.Cells["A76:D76"].Merge = true;
                    //_genericSheet.Cells["A76:D76"].Style.Locked = true;
                    //_genericSheet.Cells["A76:D76"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A76:D76"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A76:D76"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E76:G76"].Value = "1.4.7 Correo electrónico";
                    //_genericSheet.Cells["E76:G76"].Merge = true;
                    //_genericSheet.Cells["E76:G76"].Style.Locked = true;
                    //_genericSheet.Cells["E76:G76"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E76:G76"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E76:G76"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["H76:J76"].Value = "1.4.7 Correo electrónico";
                    //_genericSheet.Cells["H76:J76"].Merge = true;
                    //_genericSheet.Cells["H76:J76"].Style.Locked = true;
                    //_genericSheet.Cells["H76:J76"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H76:J76"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["H76:J76"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A77:D77"].Value = Organizacion[0].vCorreoElectronicoOrg;
                    //_genericSheet.Cells["A77:D77"].Merge = true;
                    //_genericSheet.Cells["A77:D77"].Style.Locked = true;
                    //_genericSheet.Cells["A77:D77"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A77:D77"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A77:D77"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["E77:G77"].Value = Organizacion[1].vCorreoElectronicoOrg;
                    //_genericSheet.Cells["E77:G77"].Merge = true;
                    //_genericSheet.Cells["E77:G77"].Style.Locked = true;
                    //_genericSheet.Cells["E77:G77"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["E77:G77"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E77:G77"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["H77:J77"].Value = Organizacion[2].vCorreoElectronicoOrg;
                    //_genericSheet.Cells["H77:J77"].Merge = true;
                    //_genericSheet.Cells["H77:J77"].Style.Locked = true;
                    //_genericSheet.Cells["H77:J77"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H77:J77"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["H77:J77"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A78:D78"].Value = "1.4.9 Tipo de organización";
                    //_genericSheet.Cells["A78:D78"].Merge = true;
                    //_genericSheet.Cells["A78:D78"].Style.Locked = true;
                    //_genericSheet.Cells["A78:D78"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A78:D78"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A78:D78"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E78:G78"].Value = "1.4.9 Tipo de organización";
                    //_genericSheet.Cells["E78:G78"].Merge = true;
                    //_genericSheet.Cells["E78:G78"].Style.Locked = true;
                    //_genericSheet.Cells["E78:G78"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["E78:G78"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E78:G78"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["H78:J78"].Value = "1.4.9 Tipo de organización";
                    //_genericSheet.Cells["H78:J78"].Merge = true;
                    //_genericSheet.Cells["H78:J78"].Style.Locked = true;
                    //_genericSheet.Cells["H78:J78"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H78:J78"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["H78:J78"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///**/

                    //_genericSheet.Cells["A79:D79"].Value = Organizacion[0].vOrganizacion;
                    //_genericSheet.Cells["A79:D79"].Merge = true;
                    //_genericSheet.Cells["A79:D79"].Style.Locked = true;
                    //_genericSheet.Cells["A79:D79"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["A79:D79"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["A79:D79"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["E79:G79"].Value = Organizacion[1].vOrganizacion;
                    //_genericSheet.Cells["E79:G79"].Merge = true;
                    //_genericSheet.Cells["E79:G79"].Style.Locked = true;
                    //_genericSheet.Cells["E79:G79"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["E79:G79"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["E79:G79"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;


                    //_genericSheet.Cells["H79:J79"].Value = Organizacion[2].vOrganizacion;
                    //_genericSheet.Cells["H79:J79"].Merge = true;
                    //_genericSheet.Cells["H79:J79"].Style.Locked = true;
                    //_genericSheet.Cells["H79:J79"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H79:J79"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet.Cells["H79:J79"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/
                    //#endregion

                    ///*****************************************************************************************/

                    //#region CABECERA Lista de productores

                    ///*****************************************************************************************/

                    //_genericSheet.Cells["A81:D81"].Value = "1.4.8 N° de productores agrupados que se espera atender";
                    //_genericSheet.Cells["A81:D81"].Merge = true;
                    //_genericSheet.Cells["A81:D81"].Style.Locked = true;
                    //_genericSheet.Cells["A81:D81"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["A81:D81"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["A81:D81"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["B82:B82"].Value = "N°";
                    //_genericSheet.Cells["B82:B82"].Merge = true;
                    //_genericSheet.Cells["B82:B82"].Style.Locked = true;
                    //_genericSheet.Cells["B82:B82"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet.Cells["B82:B82"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["B82:B82"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    ////_genericSheet.Cells["B82:B82"].Style.Border.;
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["C82:D82"].Value = "Apellidos y Nombres";
                    //_genericSheet.Cells["C82:D82"].Merge = true;
                    //_genericSheet.Cells["C82:D82"].Style.Locked = true;
                    //_genericSheet.Cells["C82:D82"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["C82:D82"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["C82:D82"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["E82:E82"].Value = "DNI";
                    //_genericSheet.Cells["E82:E82"].Merge = true;
                    //_genericSheet.Cells["E82:E82"].Style.Locked = true;
                    //_genericSheet.Cells["E82:E82"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["E82:E82"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["E82:E82"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["F82:F82"].Value = "Celular";
                    //_genericSheet.Cells["F82:F82"].Merge = true;
                    //_genericSheet.Cells["F82:F82"].Style.Locked = true;
                    //_genericSheet.Cells["F82:F82"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["F82:F82"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["F82:F82"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["G82:G82"].Value = "Edad";
                    //_genericSheet.Cells["G82:G82"].Merge = true;
                    //_genericSheet.Cells["G82:G82"].Style.Locked = true;
                    //_genericSheet.Cells["G82:G82"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["G82:G82"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["G82:G82"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["H82:H82"].Value = "Sexo";
                    //_genericSheet.Cells["H82:H82"].Merge = true;
                    //_genericSheet.Cells["H82:H82"].Style.Locked = true;
                    //_genericSheet.Cells["H82:H82"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["H82:H82"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["H82:H82"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    //_genericSheet.Cells["I82:I82"].Value = "Recibio capacitación y/o asistencia técnica";
                    //_genericSheet.Cells["I82:I82"].Merge = true;
                    //_genericSheet.Cells["I82:I82"].Style.Locked = true;
                    //_genericSheet.Cells["I82:I82"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet.Cells["I82:I82"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet.Rows[82].Height = 30.6;
                    //_genericSheet.Cells["I82:I82"].Style.WrapText = true;
                    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet.Cells["I82:I82"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    //_genericSheet.Column(1).Style.Locked = true;
                    //_genericSheet.Workbook.Protection.LockWindows = true;
                    //_genericSheet.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //#endregion


                    ///*****************************************************************************************/


                    ///*****************************************************************************************/

                    //#region TABLA LISTA PRODUCTORES
                    //List<Productor> Productor = costoDL.ListarProductorRpt(fichaTecnica);
                    //int rowIndexComp = 83;
                    //int colcomp = 2;
                    //double ContMas = 0;
                    //double ContFem = 0;
                    //double ContJov = 0;
                    //double contCap = 0;
                    //double sumEdad = 0;
                    //for (int i = 0; i < Productor.Count; i++)
                    //{
                    //    _texto_row1(_genericSheet, rowIndexComp, colcomp++, Productor[i].Nro, "#E2EFDA");
                    //    _texto_row1(_genericSheet, rowIndexComp, colcomp++, Productor[i].vApellidosNombres, "#E2EFDA");


                    //    //_genericSheet.Cells["C" + rowIndexComp + ":D" + rowIndexComp].Value = Productor[i].vApellidosNombres;
                    //    _genericSheet.Cells["C" + rowIndexComp + ":D" + rowIndexComp].Merge = true;
                    //    _genericSheet.Cells["C" + rowIndexComp + ":D" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet.Cells["C" + rowIndexComp + ":D" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //    //colFromHex = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    //_genericSheet.Cells["C" + rowIndexComp + ":D" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex);

                    //    //_texto_row1(_genericSheet, rowIndexComp, colcomp++, Productor[i].vApellidosNombres, "#E2EFDA");
                    //    colcomp = colcomp + 1;
                    //    _texto_row1(_genericSheet, rowIndexComp, colcomp++, Productor[i].vDni, "#E2EFDA");
                    //    _texto_row1(_genericSheet, rowIndexComp, colcomp++, Productor[i].vCelular, "#E2EFDA");
                    //    _texto_row1(_genericSheet, rowIndexComp, colcomp++, Productor[i].iEdad.ToString(), "#E2EFDA");
                    //    _texto_row1(_genericSheet, rowIndexComp, colcomp++, Productor[i].vSexo, "#E2EFDA");
                    //    _texto_row1(_genericSheet, rowIndexComp, colcomp++, Productor[i].vRecibioCapacitacion, "#E2EFDA");
                    //    double rpta = (Productor[i].vSexo == "Masculino") ? ContMas++ : ContFem++;
                    //    rpta = (Convert.ToDouble(Productor[i].iEdad) < 25) ? ContJov++ : 0;
                    //    rpta = (Productor[i].vRecibioCapacitacion == "SI") ? contCap++ : 0;

                    //    rowIndexComp++;
                    //    colcomp = 2;
                    //    sumEdad = sumEdad + Productor[i].iEdad;
                    //}

                    ////rowIndexComp--;
                    //colcomp++;
                    //colcomp++;
                    //_texto_row1(_genericSheet, rowIndexComp, colcomp++, "Total", "#FFFFFF");
                    //_texto_row1(_genericSheet, rowIndexComp, colcomp++, (ContMas + ContFem).ToString(), "#FFFFFF");
                    //_texto_row1(_genericSheet, rowIndexComp, colcomp++, "Promedio Edad", "#FFFFFF");
                    //_texto_row1(_genericSheet, rowIndexComp++, colcomp++, sumEdad / (ContMas + ContFem), "#FFFFFF");

                    //colcomp = 7;
                    //_texto_row1(_genericSheet, rowIndexComp, colcomp++, "Masculino", "#FFFFFF");
                    //_texto_row1(_genericSheet, rowIndexComp++, colcomp++, (ContMas).ToString(), "#FFFFFF");

                    //colcomp = 7;
                    //_texto_row1(_genericSheet, rowIndexComp, colcomp++, "Femenino", "#FFFFFF");
                    //_texto_row1(_genericSheet, rowIndexComp++, colcomp++, (ContFem).ToString(), "#FFFFFF");

                    //colcomp = 7;
                    //_texto_row1(_genericSheet, rowIndexComp, colcomp++, "Jovenes", "#FFFFFF");
                    //_texto_row1(_genericSheet, rowIndexComp++, colcomp++, (ContJov).ToString(), "#FFFFFF");

                    //colcomp = 7;
                    //_texto_row1(_genericSheet, rowIndexComp, colcomp++, "% Femenino", "#FFFFFF");
                    //_texto_row1(_genericSheet, rowIndexComp++, colcomp++, ((ContFem / (ContMas + ContFem)) * 100).ToString("0. ##"), "#FFFFFF");

                    //colcomp = 7;
                    //_texto_row1(_genericSheet, rowIndexComp, colcomp++, "% Jovenes", "#FFFFFF");
                    //_texto_row1(_genericSheet, rowIndexComp++, colcomp++, ((ContJov / (ContMas + ContFem)) * 100).ToString("0. ##"), "#FFFFFF");

                    //colcomp = 7;
                    //_texto_row1(_genericSheet, rowIndexComp, colcomp++, "Recibio CAP y/o AT", "#FFFFFF");
                    //_texto_row1(_genericSheet, rowIndexComp++, colcomp++, (contCap).ToString(), "#FFFFFF");

                    //colcomp = 7;
                    //_texto_row1(_genericSheet, rowIndexComp, colcomp++, "% Recibieron CAP y/o AT", "#FFFFFF");
                    //_texto_row1(_genericSheet, rowIndexComp++, colcomp++, ((contCap / (ContMas + ContFem)) * 100).ToString("0. ##"), "#FFFFFF");

                    //#endregion

                    ///****************************************************************************************/
                    //#endregion

                    //#region Identificación
                    //NombreReporte = "Identificación";
                    //var _genericSheet_ = excelPackage.Workbook.Worksheets.Add(NombreReporte);
                    //_genericSheet_.View.ShowGridLines = false;
                    //_genericSheet_.View.ZoomScale = 100;
                    //_genericSheet_.PrinterSettings.PaperSize = ePaperSize.A4;
                    //_genericSheet_.PrinterSettings.FitToPage = true;
                    //_genericSheet_.PrinterSettings.Orientation = eOrientation.Landscape;
                    //_genericSheet_.View.PageBreakView = true;

                    //#region CABECERA

                    //_genericSheet_.Cells["A2:H2"].Value = "IDENTIFICACIÓN";
                    //_genericSheet_.Cells["A2:H2"].Merge = true;
                    //_genericSheet_.Cells["A2:H2"].Style.Locked = true;
                    //_genericSheet_.Cells["A2:H2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //System.Drawing.Color colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#548235");
                    //_genericSheet_.Cells["A2:H2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["A2:H2"].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //System.Drawing.Color colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["A2:H2"].Style.Font.Size = 18;
                    //_genericSheet_.Cells["A2:H2"].Style.Font.Bold = true;
                    //_genericSheet_.Cells["A2:H2"].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["A2:H2"].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //_genericSheet_.Cells["A4:H4"].Value = "2.1 DESCRIPCIÓN DE LA SITUACIÓN ACTUAL";
                    //_genericSheet_.Cells["A4:H4"].Merge = true;
                    //_genericSheet_.Cells["A4:H4"].Style.Locked = true;
                    //_genericSheet_.Cells["A4:H4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#548235");
                    //_genericSheet_.Cells["A4:H4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["A4:H4"].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["A4:H4"].Style.Font.Size = 11;
                    //_genericSheet_.Cells["A4:H4"].Style.Font.Bold = true;
                    //_genericSheet_.Cells["A4:H4"].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["A4:H4"].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //_genericSheet_.Cells["A6:H6"].Value = "2.1.1 Rango de tecnologías usadas en las unidades productivas";
                    //_genericSheet_.Cells["A6:H6"].Merge = true;
                    //_genericSheet_.Cells["A6:H6"].Style.Locked = true;
                    //_genericSheet_.Cells["A6:H6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#F2F2F2");
                    //_genericSheet_.Cells["A6:H6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["A6:H6"].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["A6:H6"].Style.Font.Size = 11;
                    //_genericSheet_.Cells["A6:H6"].Style.Font.Bold = false;
                    //_genericSheet_.Cells["A6:H6"].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["A6:H6"].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //#endregion

                    //#region TABLA DE SITUACION ACTUAL

                    //_genericSheet_.Cells["B8:C8"].Value = "Tecnologías/prácticas utilizadas en la situación actual";
                    //_genericSheet_.Cells["B8:C8"].Merge = true;
                    //_genericSheet_.Cells["B8:C8"].Style.Locked = true;
                    //_genericSheet_.Cells["B8:C8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["B8:C8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B8:C8"].Style.WrapText = true;
                    //_genericSheet_.Cells["B8:C8"].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet_.Cells["D8:F8"].Value = "Tecnologías/prácticas presentes en el mercado";
                    //_genericSheet_.Cells["D8:F8"].Merge = true;
                    //_genericSheet_.Cells["D8:F8"].Style.Locked = true;
                    //_genericSheet_.Cells["D8:F8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["D8:F8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["D8:F8"].Style.WrapText = true;
                    //_genericSheet_.Cells["D8:F8"].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet_.Cells["G8:H8"].Value = "Tecnológia idonea para los productores (las que necesita el productor)";
                    //_genericSheet_.Cells["G8:H8"].Merge = true;
                    //_genericSheet_.Cells["G8:H8"].Style.Locked = true;
                    //_genericSheet_.Cells["G8:H8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["G8:H8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["G8:H8"].Style.WrapText = true;
                    //_genericSheet_.Cells["G8:H8"].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    /////*****************************************************************************************/

                    ////int ol = 0;
                    //#endregion

                    //#region DataTecnologias
                    //List<Tecnologias> tecnologias = costoDL.SP_Listar_TecnologiasRpt(fichaTecnica);
                    //rowIndexComp = 9;
                    //colcomp = 2;
                    //for (int i = 0; i < tecnologias.Count; i++)
                    //{
                    //    _texto_row1(_genericSheet_, rowIndexComp, colcomp++, tecnologias[i].vtecnologia1, "#E2EFDA");
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":C" + rowIndexComp].Merge = true;
                    //    colcomp++;
                    //    _texto_row1(_genericSheet_, rowIndexComp, colcomp++, tecnologias[i].vtecnologia2, "#E2EFDA");
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":F" + rowIndexComp].Merge = true;
                    //    colcomp++;
                    //    colcomp++;
                    //    _texto_row1(_genericSheet_, rowIndexComp, colcomp++, tecnologias[i].vtecnologia3, "#E2EFDA");
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":H" + rowIndexComp].Merge = true;

                    //    rowIndexComp++;
                    //    colcomp = 2;
                    //}

                    //_texto_row1(_genericSheet_, rowIndexComp, colcomp++, "Tecnologías actuales", "#FFFFFF");
                    //_texto_row1(_genericSheet_, rowIndexComp, colcomp++, tecnologias.Count, "#FFFFFF");

                    //colcomp = 7;
                    //_texto_row1(_genericSheet_, rowIndexComp, colcomp++, "Total brecha", "#FFFFFF");
                    //_texto_row1(_genericSheet_, rowIndexComp, colcomp++, tecnologias.Count, "#FFFFFF");

                    //rowIndexComp = rowIndexComp + 2;

                    //#endregion

                    //#region Limitaciones o Barreras  / Situacion Actual

                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Value = "2.1.2 Limitaciones o barreras al uso o aprovechamiento de conocimientos y/o tecnologías en las unidades productivas";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#F2F2F2");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //rowIndexComp++;
                    //_genericSheet_.Rows[rowIndexComp].Height = 124.2;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Value = "EL BAJO NIVEL EDUCATIVO DE LA POBLACION DEL SECTOR RURAL ES UNA  LIMITACION PARA EL DESARROLLO DE CAPACIDADAES DE LOS AGRICULTORES ,  LIMITA LA CAPACIDAD DE LOS PRODUCTORES PARA LA INNOVACION TECNOLOGICA Y PARA SU CAPACIDAD DE GESTION.  EL ACCESO A LA INFORMACION AGRARIA ES LIMITADO   MAS AUN EN EL MANEJO INTEGRADO DE PLAGAS EN CULTIVO DE CAFE DEBIDO A LA FALTA DE INFRAESTRUCTURA Y MEDIOS DE COMUNICACION Y A LA FALTA DE INVERSION PUBLICA EN EL MEDIO,  NO HAY FUENTES DE INFORMACION SOBRE EL MANEJO FITOSANITARIO DE LOS CULTIVOS, LA BAJA INVERSION EN LA CONDUCCION DE SUS CULTIVOS ES UNA BARRERA A USO DE TECNOLOGIAS EN SUS UNIDADES PRODUCTIVAS.";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#F2F2F2");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;
                    //int RowLimitacion = rowIndexComp;

                    //rowIndexComp++;
                    //rowIndexComp++;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Value = "2.1.3 Estado situacional de la provisión y acceso a servicios de extensión";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#F2F2F2");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //rowIndexComp++;
                    //_genericSheet_.Rows[rowIndexComp].Height = 124.2;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Value = "LOS SERVICIOS DE ASISTENCIA TECNICA , TRANSFERENCIA DE TECNOLOGIAS Y CAPACITACIONES ES MUY LIMITADO, SOBRETODO EN EL MANEJO FITOSANITARIO Y NUTRICIONAL DE LOS CULTIVOS, PUES MUCHOS DE LOS AGRICULTORES NO REALIZAN SUS PRACTICAS ADECUADAS POR DESCONOCIMIENTO  Y SI TENEN UN PROBLEMA SERIO DE PLAGAS VAN  A LAS AGROVETERINARIAS Y COMPRAN LO QUE LES DEN, SIN TENER UN MONITOREO DE CAMPO NI EL CONOCIMIENTO REAL DEL PROBLEMA. NO HAY PRESCENCIA  DEL SECTOR PRIVADO NI  DEL ESTADO QUE LES DEN ASISTENCIA TECNICA Y CAPACITACIONES LO QUE HACE QUE NO TENGAN  ACCESO A SERVCIOS DE EXTECION AGRARIA.";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#F2F2F2");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //int RowEstadoSituacional = rowIndexComp;

                    //rowIndexComp++;
                    //rowIndexComp++;
                    //#endregion

                    //#region 2.2 IDENTIFICACIÓN DEL PROBLEMA CENTRAL, CAUSAS Y EFECTOS

                    //_genericSheet_.Cells["A4:H4"].Value = "2.2 IDENTIFICACIÓN DEL PROBLEMA CENTRAL, CAUSAS Y EFECTOS";
                    //_genericSheet_.Cells["A4:H4"].Merge = true;
                    //_genericSheet_.Cells["A4:H4"].Style.Locked = true;
                    //_genericSheet_.Cells["A4:H4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#548235");
                    //_genericSheet_.Cells["A4:H4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["A4:H4"].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["A4:H4"].Style.Font.Size = 11;
                    //_genericSheet_.Cells["A4:H4"].Style.Font.Bold = true;
                    //_genericSheet_.Cells["A4:H4"].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["A4:H4"].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //rowIndexComp++;
                    //rowIndexComp++;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Value = "2.2.1 Definición y caracterización del problema central";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#F2F2F2");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;
                    //rowIndexComp++;

                    ////rowIndexComp++;

                    //#endregion

                    //#region TABLA DE CAUSAS/EFECTOS

                    //_genericSheet_.Rows[rowIndexComp].Height = 57.6;

                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Value = "Causas indirectas";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Value = "Causas directas";
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Value = "Problema central";
                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Value = "Efectos directos";
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    ///*****************************************************************************************/

                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Value = "Efectos indirectos";
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //#endregion

                    //#region Data Causas / Efectos
                    //List<CausasIndirectas> causas = costoDL.SP_Listar_CausasDirectasIndirectasRpt(fichaTecnica);
                    ////rowIndexComp = 9;
                    //rowIndexComp++;
                    //int rowCab = rowIndexComp;
                    //colcomp = 2;

                    //for (int i = 0; i < causas.Count; i++)
                    //{
                    //    _texto_row1(_genericSheet_, rowIndexComp, colcomp++, causas[i].vDescrCausaInDirecta, "#E2EFDA");
                    //    _texto_row1(_genericSheet_, rowIndexComp, colcomp++, causas[i].vDescrCausaDirecta, "#E2EFDA");

                    //    _genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Value = causas[i].vProblemaCentral;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    rowIndexComp++;
                    //    colcomp = 2;
                    //}

                    //List<EfectoIndirecto> efecto = costoDL.SP_Listar_EfectosDirectosIndirectosRpt(fichaTecnica);
                    //colcomp = 6;
                    //for (int i = 0; i < efecto.Count; i++)
                    //{
                    //    _texto_row1(_genericSheet_, rowCab, colcomp++, efecto[i].vDescEfectoDirecto, "#E2EFDA");
                    //    _texto_row1(_genericSheet_, rowCab, colcomp++, efecto[i].vDescEfectoIndirecto, "#E2EFDA");

                    //    rowCab++;
                    //    colcomp = 6;
                    //}

                    //rowIndexComp++;

                    //#endregion

                    //#region 2.3 POBLACIÓN OBJETIVO

                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Value = "2.3 POBLACIÓN OBJETIVO";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#548235");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //rowIndexComp++;
                    //rowIndexComp++;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Value = "2.3.1 Población afectada en el subsector de agricultura / ganadería";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#F2F2F2");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;
                    //rowIndexComp++;
                    //rowIndexComp++;

                    ////rowIndexComp++;

                    //#endregion

                    //#region TABLA POBLACION OBJETIVO
                    //int rowIndexCompFilaSig = rowIndexComp + 1;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Value = "Número de unidades productivas de la cadena productiva beneficiarias del servicio";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.Font.Bold = false;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.WrapText = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexCompFilaSig].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Rows[rowIndexComp].Height = 42;
                    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;


                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Value = "Unidad de medida de las unidades productivas de la cadena productiva beneficiarias del servicio";
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Merge = true;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.Locked = true;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.Font.Size = 11;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.Font.Bold = false;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.WrapText = true;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexCompFilaSig].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Value = "Número de familias beneficiarias";
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Merge = true;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.Locked = true;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.Font.Size = 11;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.Font.Bold = false;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.WrapText = true;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexCompFilaSig].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Value = "Volumen de producción de la cadena productiva";
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Value = "Cantidad";
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Merge = true;
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.Locked = true;
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.Font.Size = 11;
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.Font.Bold = false;
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.WrapText = true;
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["E" + rowIndexCompFilaSig + ":E" + rowIndexCompFilaSig].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Value = "Unidad Medida";
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Merge = true;
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.Locked = true;
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.Font.Size = 11;
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.Font.Bold = false;
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.WrapText = true;
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["F" + rowIndexCompFilaSig + ":F" + rowIndexCompFilaSig].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Value = "Rendimiento promedio de la cadena productiva";
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Merge = true;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.Locked = true;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.Font.Size = 11;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.Font.Bold = false;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.WrapText = true;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexCompFilaSig].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Value = "Gremios organizacionales a las que está vinculada la cadena productiva";
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Merge = true;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.Locked = true;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.Font.Size = 11;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.Font.Bold = false;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.WrapText = true;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexCompFilaSig].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //rowIndexComp++;
                    //rowIndexComp++;

                    //#endregion

                    //#region DATA POBLACION OBJETIVO

                    //List<Identificacion> PoblacionObjs = costoDL.PA_ListarDetallePobObj_Rpt(fichaTecnica);
                    //for (int i = 0; i < PoblacionObjs.Count; i++)
                    //{
                    //    //RowLimitacion
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Value = PoblacionObjs[0].vLimitaciones;
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Merge = true;
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.Locked = true;
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.WrapText = true;
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + RowLimitacion + ":H" + RowLimitacion].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 42;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    //RowEstadoSituacional
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Value = PoblacionObjs[0].vLimitaciones;
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Merge = true;
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.Locked = true;
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.WrapText = true;
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + RowEstadoSituacional + ":H" + RowEstadoSituacional].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 42;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Value = PoblacionObjs[i].vNumeroUnidadesProductivas;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Rows[rowIndexComp].Height = 15;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Value = PoblacionObjs[i].vUnidadMedidaProductivas;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":C" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 42;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Value = PoblacionObjs[i].vNumerosFamiliares;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["D" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 42;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Value = PoblacionObjs[i].vCantidad;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 42;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Value = PoblacionObjs[i].vUnidadMedida;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 42;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Value = PoblacionObjs[i].vRendimientoCadenaProductiva;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 42;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Value = PoblacionObjs[i].vGremios;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 42;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    rowIndexComp++;
                    //    rowIndexComp++;
                    //    rowIndexComp++;

                    //}

                    //#endregion

                    //#region DEFINICION DEL OBJETIVO DEL SERVICIO

                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Value = "2.4 DEFINICIÓN DEL OBJETIVO DEL SERVICIO";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#548235");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //rowIndexComp++;
                    //rowIndexComp++;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Value = "2.4.1 Objetivo";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#F2F2F2");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Rows[6].Height = 23.4;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;
                    //rowIndexComp++;
                    //rowIndexComp++;

                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Value = "Descripción del objetivo central";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    ////_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Value = "Indicador";
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ////_genericSheet_.Rows[rowIndexComp].Height = 42;
                    ////_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;


                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Value = "Unidad de Medida";
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ////_genericSheet_.Rows[rowIndexComp].Height = 42;
                    ////_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Value = "Meta";
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ////_genericSheet_.Rows[rowIndexComp].Height = 42;
                    ////_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;

                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Value = "Medio de Verificación";
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ////_genericSheet_.Rows[rowIndexComp].Height = 42;
                    ////_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;
                    //rowIndexComp++;
                    ////rowIndexComp++;
                    //#endregion

                    //#region DATA OBJETIVO SERVICIO

                    //List<Indicadores> indicadores = costoDL.PA_ListarIndicadores_Rpt(fichaTecnica);

                    //int rowObjCentralIni = rowIndexComp;

                    //for (int i = 0; i < indicadores.Count; i++)
                    //{
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Value = indicadores[i].TipoIndicador;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Value = indicadores[i].vUnidadMedida;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Value = indicadores[i].vMeta;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Value = indicadores[i].vMedioVerificacion;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Rows[rowIndexComp].Height = 58.2;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    rowIndexComp++;

                    //}

                    //rowIndexComp--;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Value = PoblacionObjs[0].vObjetivoCentral;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Merge = true;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Locked = true;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Font.Size = 11;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Font.Bold = false;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Font.Name = "Calibri";
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.WrapText = true;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //_genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ////_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    ////_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //_genericSheet_.Column(1).Style.Locked = true;
                    //_genericSheet_.Workbook.Protection.LockWindows = true;
                    //_genericSheet_.Workbook.Protection.LockStructure = true;
                    //rowIndexComp++;
                    //#endregion

                    //#region LISTA DE COMPONENTES

                    //List<Componente> componentes = costoDL.PA_ListarComponentes_Rpt1(fichaTecnica);
                    //List<Componente> componentesDistinct = componentes;
                    //componentesDistinct = componentesDistinct.DistinctBy(z => z.nTipoComponente).ToList();
                    //for (int i = 0; i < componentesDistinct.Count(); i++)
                    //{
                    //    fichaTecnica.TotalDias = componentesDistinct[i].nTipoComponente;
                    //    List<Componente> IndicadorPorComponente = costoDL.PA_ListarComponentes_Rpt(fichaTecnica);
                    //    #region CABECERA
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Value = "Descripción del Componente";
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":D" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Rows[rowIndexComp].Height = 58.2;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Value = "Indicador";
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Value = "Unidad de Medida";
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;


                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Value = "Meta";
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Value = "Medio de Verificación";
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    rowIndexComp++;

                    //    #endregion



                    //    #region INDICADOR_COMPONENTE
                    //    rowObjCentralIni = rowIndexComp;
                    //    for (int j = 0; j < IndicadorPorComponente.Count; j++)
                    //    {
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Value = IndicadorPorComponente[j].vIndicador;
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Merge = true;
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Locked = true;
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //        colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //        colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Size = 11;
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Bold = false;
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Name = "Calibri";
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.WrapText = true;
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["E" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Rows[rowIndexComp].Height = 87;
                    //        //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //        _genericSheet_.Column(1).Style.Locked = true;
                    //        _genericSheet_.Workbook.Protection.LockWindows = true;
                    //        _genericSheet_.Workbook.Protection.LockStructure = true;

                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Value = IndicadorPorComponente[j].vUnidadMedida;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Merge = true;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Locked = true;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //        colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //        colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Size = 11;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Bold = false;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Name = "Calibri";
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.WrapText = true;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //        //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //        //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //        _genericSheet_.Column(1).Style.Locked = true;
                    //        _genericSheet_.Workbook.Protection.LockWindows = true;
                    //        _genericSheet_.Workbook.Protection.LockStructure = true;


                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Value = IndicadorPorComponente[j].vMeta;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Merge = true;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Locked = true;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //        colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //        colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Size = 11;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Bold = false;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Name = "Calibri";
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.WrapText = true;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //        //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //        //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //        _genericSheet_.Column(1).Style.Locked = true;
                    //        _genericSheet_.Workbook.Protection.LockWindows = true;
                    //        _genericSheet_.Workbook.Protection.LockStructure = true;

                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Value = IndicadorPorComponente[j].vMedio;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //        colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //        colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.WrapText = true;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //        //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //        //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //        _genericSheet_.Column(1).Style.Locked = true;
                    //        _genericSheet_.Workbook.Protection.LockWindows = true;
                    //        _genericSheet_.Workbook.Protection.LockStructure = true;

                    //        rowIndexComp++;
                    //    }

                    //    if (IndicadorPorComponente.Count > 0)
                    //    {
                    //        rowIndexComp--;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Value = componentesDistinct[i].vDescripcion;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Merge = true;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Locked = true;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //        colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //        colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Font.Size = 11;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Font.Bold = false;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Font.Name = "Calibri";
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.WrapText = true;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["B" + rowObjCentralIni + ":D" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //        //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //        //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //        _genericSheet_.Column(1).Style.Locked = true;
                    //        _genericSheet_.Workbook.Protection.LockWindows = true;
                    //        _genericSheet_.Workbook.Protection.LockStructure = true;
                    //    }


                    //    rowIndexComp++;

                    //    #region CABECERA_ACTIVIDADES

                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Value = "Actividades";
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Rows[rowIndexComp].Height = 29.5;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Value = "Descripción";
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Value = "Unidad de Medida";
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;


                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Value = "Meta";
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Value = "Medio de Verificación";
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //    colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.WrapText = true;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //    _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //    //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //    //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //    _genericSheet_.Column(1).Style.Locked = true;
                    //    _genericSheet_.Workbook.Protection.LockWindows = true;
                    //    _genericSheet_.Workbook.Protection.LockStructure = true;

                    //    rowIndexComp++;

                    //    #endregion

                    //    #region DATA_ACTIVIDADES
                    //    fichaTecnica.iCodConvocatoria = componentesDistinct[i].nTipoComponente;
                    //    List<Actividad> actividades = costoDL.ListarActividadesRpt(fichaTecnica);
                    //    for (int y = 0; y < actividades.Count; y++)
                    //    {
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Value = actividades[y].vActividad;
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Merge = true;
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Locked = true;
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //        colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //        colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Size = 11;
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Bold = false;
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Name = "Calibri";
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.WrapText = true;
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["B" + rowIndexComp + ":B" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Rows[rowIndexComp].Height = 87;
                    //        //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //        _genericSheet_.Column(1).Style.Locked = true;
                    //        _genericSheet_.Workbook.Protection.LockWindows = true;
                    //        _genericSheet_.Workbook.Protection.LockStructure = true;

                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Value = actividades[y].vDescripcion;
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Merge = true;
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Locked = true;
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //        colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //        colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Size = 11;
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Bold = false;
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Name = "Calibri";
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.WrapText = true;
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["C" + rowIndexComp + ":E" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //        //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //        //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //        _genericSheet_.Column(1).Style.Locked = true;
                    //        _genericSheet_.Workbook.Protection.LockWindows = true;
                    //        _genericSheet_.Workbook.Protection.LockStructure = true;


                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Value = actividades[y].vUnidadMedida;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Merge = true;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Locked = true;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //        colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //        colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Size = 11;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Bold = false;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Name = "Calibri";
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.WrapText = true;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["F" + rowIndexComp + ":F" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //        //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //        //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //        _genericSheet_.Column(1).Style.Locked = true;
                    //        _genericSheet_.Workbook.Protection.LockWindows = true;
                    //        _genericSheet_.Workbook.Protection.LockStructure = true;

                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Value = actividades[y].vMeta;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Merge = true;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Locked = true;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //        colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //        colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Size = 11;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Bold = false;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Name = "Calibri";
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.WrapText = true;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["G" + rowIndexComp + ":G" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //        //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //        //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //        _genericSheet_.Column(1).Style.Locked = true;
                    //        _genericSheet_.Workbook.Protection.LockWindows = true;
                    //        _genericSheet_.Workbook.Protection.LockStructure = true;


                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Value = actividades[y].vMedio;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Merge = true;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Locked = true;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //        colFromHex_ = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Fill.BackgroundColor.SetColor(colFromHex_);
                    //        colFromHex1_ = System.Drawing.ColorTranslator.FromHtml("#000000");
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Size = 11;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Bold = false;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Name = "Calibri";
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.WrapText = true;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Font.Color.SetColor(colFromHex1_);
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //        _genericSheet_.Cells["H" + rowIndexComp + ":H" + rowIndexComp].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //        //_genericSheet_.Rows[rowIndexComp].Height = 29.4;
                    //        //_genericSheet_.Rows[rowIndexComp + 1].Height = 73.2;
                    //        _genericSheet_.Column(1).Style.Locked = true;
                    //        _genericSheet_.Workbook.Protection.LockWindows = true;
                    //        _genericSheet_.Workbook.Protection.LockStructure = true;

                    //        rowIndexComp++;
                    //    }
                    //    #endregion


                    //    #endregion


                    //}




                    //#endregion

                    //#endregion

                    //#region Cronograma
                    //NombreReporte = "Cronograma";
                    //var _genericSheet___ = excelPackage.Workbook.Worksheets.Add(NombreReporte);
                    //_genericSheet___.View.ShowGridLines = false;
                    //_genericSheet___.View.ZoomScale = 100;
                    //_genericSheet___.PrinterSettings.PaperSize = ePaperSize.A4;
                    //_genericSheet___.PrinterSettings.FitToPage = true;
                    //_genericSheet___.PrinterSettings.Orientation = eOrientation.Landscape;
                    //_genericSheet___.View.PageBreakView = true;


                    //List<string> _cabecera = new List<string>();

                    //_cabecera.Add("Nro");
                    //_cabecera.Add("Actividades");
                    //_cabecera.Add("Descripcion");
                    //_cabecera.Add("Unidad Medida");
                    //_cabecera.Add("Meta");

                    //int finish = pintarcabeceras(_cabecera, _genericSheet___, "CRONOGRAMA DETALLADO");

                    //_genericSheet___.Cells["A3:D3"].Value = "3.1 CRONOGRAMA DE EJECUCIÓN";
                    //_genericSheet___.Cells["A3:D3"].Merge = true;
                    //_genericSheet___.Cells["A3:D3"].Style.Locked = true;
                    //System.Drawing.Color colFromHex_cro = System.Drawing.ColorTranslator.FromHtml("#72aea5");
                    //_genericSheet___.Cells["A3:D3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet___.Cells["A3:D3"].Style.Fill.BackgroundColor.SetColor(colFromHex_cro);
                    //_genericSheet___.Column(1).Style.Locked = true;
                    //_genericSheet___.Workbook.Protection.LockWindows = true;
                    //_genericSheet___.Workbook.Protection.LockStructure = true;

                    ////pintar componente 1

                    ////List<Componente> componentes = cronogramaDL.ListarComponentes(cronograma);
                    //Cronograma cronograma1 = new Cronograma();
                    //cronograma1.iCodExtensionista = fichaTecnica.iCodExtensionista;
                    //DataTable componente_cro = cronogramaDL.ListarComponentesRpt(cronograma1);
                    //rowIndexComp = 8;
                    //colcomp = 1;

                    //if (componente_cro.Rows.Count > 0)
                    //{
                    //    for (int i = 0; i < componente_cro.Rows.Count; i++)
                    //    {
                    //        // Componentes
                    //        _texto_row(_genericSheet___, rowIndexComp, colcomp++, (i + 1).ToString(), "#fff2cc");
                    //        _texto_row(_genericSheet___, rowIndexComp, colcomp++, componente_cro.Rows[i][2], "#fff2cc");
                    //        _texto_row(_genericSheet___, rowIndexComp, colcomp++, componente_cro.Rows[i][5], "#fff2cc");
                    //        _texto_row(_genericSheet___, rowIndexComp, colcomp++, componente_cro.Rows[i][3], "#fff2cc");
                    //        _texto_row(_genericSheet___, rowIndexComp, colcomp++, int.Parse(componente_cro.Rows[i][6].ToString()), "#fff2cc");

                    //        for (int j = 9; j < componente_cro.Columns.Count; j++)
                    //        {
                    //            _texto_row_fecha(_genericSheet___, 6, colcomp, Convert.ToDateTime(componente_cro.Columns[j].ColumnName), "#b4c6e7");
                    //            _genericSheet___.Rows[6].Height = 40;
                    //            _texto_row1(_genericSheet___, 7, colcomp, j - 8, "#b4c6e7");
                    //            _texto_row1(_genericSheet___, rowIndexComp, colcomp, componente_cro.Rows[i][j].ToString() == "" ? "-" : componente_cro.Rows[i][j].ToString(), "#fff2cc");
                    //            colcomp++;
                    //        }
                    //        rowIndexComp++;
                    //        colcomp = 1;
                    //        //------------------------------------------------------------------------------------------------------------------------------------------------------
                    //        // Actividades
                    //        Actividad actividad_cro = new Actividad();
                    //        actividad_cro.iCodIdentificacion = Convert.ToInt32(componente_cro.Rows[i][0].ToString());
                    //        actividad_cro.iCodExtensionista = cronograma1.iCodExtensionista;
                    //        DataTable listaactividades = cronogramaDL.ListarActividadesPorComponente(actividad_cro);
                    //        for (int k = 0; k < listaactividades.Rows.Count; k++)
                    //        {
                    //            _texto_row(_genericSheet___, rowIndexComp, colcomp++, (i + 1).ToString() + "." + (k + 1).ToString(), "#FFFFFF");
                    //            _texto_row(_genericSheet___, rowIndexComp, colcomp++, listaactividades.Rows[k][3], "#FFFFFF");
                    //            _texto_row(_genericSheet___, rowIndexComp, colcomp++, listaactividades.Rows[k][4], "#FFFFFF");
                    //            _texto_row(_genericSheet___, rowIndexComp, colcomp++, listaactividades.Rows[k][5], "#FFFFFF");
                    //            _texto_row(_genericSheet___, rowIndexComp, colcomp++, int.Parse(listaactividades.Rows[k][6].ToString()), "#FFFFFF");
                    //            for (int z = 7; z < listaactividades.Columns.Count; z++)
                    //            {
                    //                if (listaactividades.Rows[k][z].ToString() == "")
                    //                {
                    //                    _texto_row1(_genericSheet___, rowIndexComp, colcomp, listaactividades.Rows[k][z].ToString(), "#E2EFDA");
                    //                }
                    //                else
                    //                {
                    //                    _texto_row1(_genericSheet___, rowIndexComp, colcomp, listaactividades.Rows[k][z].ToString(), "#00B050");
                    //                }

                    //                colcomp++;
                    //            }
                    //            rowIndexComp++;
                    //            colcomp = 1;
                    //        }
                    //    }
                    //    int totalmetas = 0;
                    //}

                    //#endregion

                    //#region Costos
                    ///*************************************************************************************/
                    //NombreReporte = "Costos";
                    //var _genericSheet__ = excelPackage.Workbook.Worksheets.Add(NombreReporte);
                    //_genericSheet__.View.ShowGridLines = false;
                    //_genericSheet__.View.ZoomScale = 100;
                    //_genericSheet__.PrinterSettings.PaperSize = ePaperSize.A4;
                    //_genericSheet__.PrinterSettings.FitToPage = true;
                    //_genericSheet__.PrinterSettings.Orientation = eOrientation.Landscape;
                    //_genericSheet__.View.PageBreakView = true;


                    //List<string> _cabecera__ = new List<string>();

                    //_cabecera__.Add("Nro");
                    //_cabecera__.Add("Actividades");
                    //_cabecera__.Add("Descripcion");
                    //_cabecera__.Add("Unidad Medida");
                    //_cabecera__.Add("Cantidad");
                    //_cabecera__.Add("Costo Unitario");
                    //_cabecera__.Add("SubTotal");

                    //finish = pintarcabeceras(_cabecera__, _genericSheet__, "COSTO");

                    //_genericSheet__.Cells["A3:D3"].Value = "3.1 COSTO DE INVERSIÓN POR COMPONENTE / ACTIVIDAD / GASTO ELEGIBLE";
                    //_genericSheet__.Cells["A3:D3"].Merge = true;
                    //_genericSheet__.Cells["A3:D3"].Style.Locked = true;
                    //System.Drawing.Color colFromHex_Costo = System.Drawing.ColorTranslator.FromHtml("#72aea5");
                    //_genericSheet__.Cells["A3:D3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //_genericSheet__.Cells["A3:D3"].Style.Fill.BackgroundColor.SetColor(colFromHex_Costo);
                    //_genericSheet__.Column(1).Style.Locked = true;
                    //_genericSheet__.Workbook.Protection.LockWindows = true;
                    //_genericSheet__.Workbook.Protection.LockStructure = true;

                    ////pintar componente 1

                    ////List<Componente> componentes = cronogramaDL.ListarComponentes(cronograma);
                    //Cronograma cronograma = new Cronograma();
                    //cronograma.iCodExtensionista = fichaTecnica.iCodExtensionista;
                    //DataTable componentes_ = costoDL.ListarComponentesRpt(cronograma);
                    //rowIndexComp = 8;
                    //colcomp = 1;

                    //if (componentes_.Rows.Count > 0)
                    //{
                    //    for (int i = 0; i < componentes_.Rows.Count; i++)
                    //    {
                    //        //------------------------------------------------------------------------------------------------------------------------------------------------------
                    //        // Componentes
                    //        _texto_row(_genericSheet__, rowIndexComp, colcomp++, (i + 1).ToString(), "#fff2cc");
                    //        _texto_row(_genericSheet__, rowIndexComp, colcomp++, componentes_.Rows[i][2], "#fff2cc");
                    //        _texto_row(_genericSheet__, rowIndexComp, colcomp++, componentes_.Rows[i][6], "#fff2cc");
                    //        _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#fff2cc");
                    //        _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#fff2cc");
                    //        _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#fff2cc");
                    //        _texto_row(_genericSheet__, rowIndexComp, colcomp++, componentes_.Rows[i][5], "#fff2cc");

                    //        for (int j = 10; j < componentes_.Columns.Count; j++)
                    //        {
                    //            _texto_row_fecha(_genericSheet__, 6, colcomp, Convert.ToDateTime(componentes_.Columns[j].ColumnName), "#b4c6e7");
                    //            //_genericSheet.Rows[6].Height = 40;
                    //            _texto_row1(_genericSheet__, 7, colcomp, j - 9, "#b4c6e7");
                    //            _texto_row1(_genericSheet__, rowIndexComp, colcomp, componentes_.Rows[i][j].ToString() == "" ? "-" : componentes_.Rows[i][j].ToString(), "#fff2cc");
                    //            colcomp++;
                    //        }
                    //        rowIndexComp++;
                    //        colcomp = 1;
                    //        //------------------------------------------------------------------------------------------------------------------------------------------------------
                    //        // Actividades
                    //        Actividad actividad_costo = new Actividad();
                    //        actividad_costo.iCodIdentificacion = Convert.ToInt32(componentes_.Rows[i][0].ToString());
                    //        actividad_costo.iCodExtensionista = cronograma.iCodExtensionista;
                    //        DataTable listaactividades = costoDL.ListarActividadesPorComponente(actividad_costo);
                    //        for (int k = 0; k < listaactividades.Rows.Count; k++)
                    //        {
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, (i + 1).ToString() + "." + (k + 1).ToString(), "#FFFFFF");
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaactividades.Rows[k][2], "#FFFFFF");
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaactividades.Rows[k][3], "#FFFFFF");
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaactividades.Rows[k][4], "#FFFFFF");
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaactividades.Rows[k][5], "#FFFFFF");
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#FFFFFF");
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaactividades.Rows[k][6], "#FFFFFF");
                    //            for (int z = 7; z < listaactividades.Columns.Count; z++)
                    //            {
                    //                //_texto_row_fecha(_genericSheet, 6, colcomp, Convert.ToDateTime(componentes.Columns[j].ColumnName), "#b4c6e7");
                    //                //_genericSheet.Rows[6].Height = 40;
                    //                //_texto_row1(_genericSheet, 7, colcomp, j - 8, "#b4c6e7");
                    //                if (listaactividades.Rows[k][z].ToString() == "")
                    //                {
                    //                    _texto_row1(_genericSheet__, rowIndexComp, colcomp, listaactividades.Rows[k][z].ToString(), "#E2EFDA");
                    //                }
                    //                else
                    //                {
                    //                    _texto_row1(_genericSheet__, rowIndexComp, colcomp, listaactividades.Rows[k][z].ToString(), "#00B050");
                    //                }

                    //                colcomp++;
                    //            }
                    //            rowIndexComp++;
                    //            colcomp = 1;
                    //            //------------------------------------------------------------------------------------------------------------------------------------------------------
                    //            //Costos Por Actividad
                    //            Actividad actividad_costo1 = new Actividad();
                    //            actividad_costo1.iCodActividad = Convert.ToInt32(listaactividades.Rows[k][0].ToString());
                    //            actividad_costo1.iCodExtensionista = cronograma.iCodExtensionista;
                    //            DataTable listaCostos = costoDL.ListarCostosPorActividad(actividad_costo1);
                    //            for (int m = 0; m < listaCostos.Rows.Count; m++)
                    //            {
                    //                _texto_row(_genericSheet__, rowIndexComp, colcomp++, (i + 1).ToString() + "." + (k + 1).ToString() + "." + (m + 1).ToString(), "#FFFFFF");
                    //                _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaCostos.Rows[m][1], "#FFFFFF");
                    //                _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaCostos.Rows[m][2], "#FFFFFF");
                    //                _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaCostos.Rows[m][3], "#FFFFFF");
                    //                _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaCostos.Rows[m][4], "#FFFFFF");
                    //                _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaCostos.Rows[m][5], "#FFFFFF");
                    //                _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaCostos.Rows[m][6], "#FFFFFF");
                    //                //_texto_row(_genericSheet, rowIndexComp, colcomp++, listaCostos.Rows[m][6], "#FFFFFF");
                    //                for (int n = 7; n < listaCostos.Columns.Count; n++)
                    //                {
                    //                    //_texto_row_fecha(_genericSheet, 6, colcomp, Convert.ToDateTime(componentes.Columns[j].ColumnName), "#b4c6e7");
                    //                    //_genericSheet.Rows[6].Height = 40;
                    //                    //_texto_row1(_genericSheet, 7, colcomp, j - 8, "#b4c6e7");
                    //                    if (listaCostos.Rows[m][n].ToString() == "")
                    //                    {
                    //                        _texto_row1(_genericSheet__, rowIndexComp, colcomp, listaCostos.Rows[m][n].ToString(), "#E2EFDA");
                    //                    }
                    //                    else
                    //                    {
                    //                        _texto_row1(_genericSheet__, rowIndexComp, colcomp, listaCostos.Rows[m][n].ToString(), "#00B050");
                    //                    }

                    //                    colcomp++;
                    //                }
                    //                rowIndexComp++;
                    //                colcomp = 1;
                    //            }

                    //        }

                    //    }


                    //    // Suma de Totales de Meta de Compromisos
                    //    int totalmetas = 0;

                    //    //foreach (Componente item in componentes)
                    //    //{
                    //    //    totalmetas = totalmetas + int.Parse(item.vMeta);
                    //    //}
                    //}

                    //rowIndexComp++;
                    //rowIndexComp++;
                    //int indCabecera = 0;

                    //if (componentes_.Rows.Count > 0)
                    //{
                    //    //_cabecera = new List<string>();

                    //    //_cabecera.Add("Nro");
                    //    //_cabecera.Add("Gasto elegible");
                    //    //_cabecera.Add("Subtotal");
                    //    ////b4c6e7
                    //    //finish = pintarcabeceras(_cabecera, _genericSheet, "COSTO");

                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "Nro", "#b4c6e7");
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "Gasto elegible", "#b4c6e7");
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#b4c6e7");
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#b4c6e7");
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#b4c6e7");
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#b4c6e7");
                    //    _genericSheet__.Cells["B" + rowIndexComp.ToString() + ":F" + rowIndexComp.ToString()].Merge = true;
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "Subtotal", "#b4c6e7");
                    //    rowIndexComp++;
                    //    colcomp = 1;
                    //    for (int i = 0; i < componentes_.Rows.Count; i++)
                    //    {

                    //        _texto_row(_genericSheet__, rowIndexComp, colcomp++, (i + 1).ToString(), "#fff2cc");
                    //        _texto_row(_genericSheet__, rowIndexComp, colcomp++, componentes_.Rows[i][6], "#fff2cc");
                    //        _genericSheet__.Cells["B" + rowIndexComp.ToString() + ":F" + rowIndexComp.ToString()].Merge = true;
                    //        colcomp = 7;
                    //        _texto_row(_genericSheet__, rowIndexComp, colcomp++, componentes_.Rows[i][5], "#fff2cc");


                    //        for (int j = 10; j < componentes_.Columns.Count; j++)
                    //        {
                    //            if (indCabecera == 0)
                    //            {
                    //                _texto_row_fecha(_genericSheet__, rowIndexComp - 1, colcomp, Convert.ToDateTime(componentes_.Columns[j].ColumnName), "#b4c6e7");
                    //            }
                    //            _texto_row1(_genericSheet__, rowIndexComp, colcomp, componentes_.Rows[i][j].ToString() == "" ? "-" : componentes_.Rows[i][j].ToString(), "#fff2cc");
                    //            colcomp++;
                    //        }
                    //        indCabecera = 1;
                    //        rowIndexComp++;
                    //        colcomp = 1;

                    //        //------------------------------------------------------------------------------------------------------------------------------------------------------
                    //        // Gastos elegibles

                    //        for (int k = 1; k < 4; k++)
                    //        {
                    //            Actividad actividad_costo1 = new Actividad();
                    //            actividad_costo1.iCodIdentificacion = Convert.ToInt32(componentes_.Rows[i][0].ToString());
                    //            actividad_costo1.iopcion = k;
                    //            actividad_costo1.iCodExtensionista = cronograma.iCodExtensionista;
                    //            DataTable listaGastosElegibles = costoDL.ListarCostosResumenPorComponente(actividad_costo1);
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, componentes_.Rows[i][6], "#FFFFFF");
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaGastosElegibles.Rows[0][0], "#FFFFFF");
                    //            _genericSheet__.Cells["B" + rowIndexComp.ToString() + ":F" + rowIndexComp.ToString()].Merge = true;
                    //            colcomp = 7;
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaGastosElegibles.Rows[0][1], "#FFFFFF");

                    //            for (int z = 2; z < listaGastosElegibles.Columns.Count; z++)
                    //            {
                    //                //_texto_row_fecha(_genericSheet, 6, colcomp, Convert.ToDateTime(componentes.Columns[j].ColumnName), "#b4c6e7");
                    //                //_genericSheet.Rows[6].Height = 40;
                    //                //_texto_row1(_genericSheet, 7, colcomp, j - 8, "#b4c6e7");
                    //                if (listaGastosElegibles.Rows[0][z].ToString() == "" || listaGastosElegibles.Rows[0][z].ToString() == "0")
                    //                {
                    //                    _texto_row1(_genericSheet__, rowIndexComp, colcomp, "-", "#E2EFDA");
                    //                }
                    //                else
                    //                {
                    //                    _texto_row1(_genericSheet__, rowIndexComp, colcomp, listaGastosElegibles.Rows[0][z].ToString(), "#00B050");
                    //                }

                    //                colcomp++;
                    //            }
                    //            rowIndexComp++;
                    //            colcomp = 1;

                    //        }



                    //    }
                    //}

                    //rowIndexComp++;
                    //rowIndexComp++;
                    //indCabecera = 0;

                    //if (componentes_.Rows.Count > 0)
                    //{
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "Nro", "#b4c6e7");
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "Gasto elegible", "#b4c6e7");
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#b4c6e7");
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#b4c6e7");
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#b4c6e7");
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "", "#b4c6e7");
                    //    _genericSheet__.Cells["B" + rowIndexComp.ToString() + ":F" + rowIndexComp.ToString()].Merge = true;
                    //    _texto_row(_genericSheet__, rowIndexComp, colcomp++, "Subtotal", "#b4c6e7");
                    //    rowIndexComp++;
                    //    //colcomp = 1;
                    //    for (int i = 0; i < 1; i++)
                    //    {


                    //        for (int j = 10; j < componentes_.Columns.Count; j++)
                    //        {
                    //            if (indCabecera == 0)
                    //            {
                    //                _texto_row_fecha(_genericSheet__, rowIndexComp - 1, colcomp, Convert.ToDateTime(componentes_.Columns[j].ColumnName), "#b4c6e7");
                    //            }
                    //            colcomp++;
                    //        }
                    //        indCabecera = 1;
                    //        colcomp = 1;

                    //        //------------------------------------------------------------------------------------------------------------------------------------------------------
                    //        // Gastos elegibles

                    //        for (int k = 1; k < 4; k++)
                    //        {
                    //            Actividad actividad_costo3 = new Actividad();
                    //            actividad_costo3.iopcion = k;
                    //            actividad_costo3.iCodExtensionista = cronograma.iCodExtensionista;
                    //            DataTable listaGastosElegibles = costoDL.ListarCostosResumenGeneral(actividad_costo3);
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, k, "#FFFFFF");
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaGastosElegibles.Rows[0][0], "#FFFFFF");
                    //            _genericSheet__.Cells["B" + rowIndexComp.ToString() + ":F" + rowIndexComp.ToString()].Merge = true;
                    //            colcomp = 7;
                    //            _texto_row(_genericSheet__, rowIndexComp, colcomp++, listaGastosElegibles.Rows[0][1], "#FFFFFF");

                    //            for (int z = 2; z < listaGastosElegibles.Columns.Count; z++)
                    //            {
                    //                if (listaGastosElegibles.Rows[0][z].ToString() == "" || listaGastosElegibles.Rows[0][z].ToString() == "0")
                    //                {
                    //                    _texto_row1(_genericSheet__, rowIndexComp, colcomp, "-", "#E2EFDA");
                    //                }
                    //                else
                    //                {
                    //                    _texto_row1(_genericSheet__, rowIndexComp, colcomp, listaGastosElegibles.Rows[0][z].ToString(), "#00B050");
                    //                }

                    //                colcomp++;
                    //            }
                    //            rowIndexComp++;
                    //            colcomp = 1;

                    //        }



                    //    }
                    //}

                    ///******************************************************************************************/

                    //#endregion

                    //#region PlanCapacitacion
                    //NombreReporte = "Plan Capacitacion";
                    //var _genericSheet_PlaCapa = excelPackage.Workbook.Worksheets.Add(NombreReporte);
                    //_genericSheet_PlaCapa.View.ShowGridLines = false;
                    //_genericSheet_PlaCapa.View.ZoomScale = 100;
                    //_genericSheet_PlaCapa.PrinterSettings.PaperSize = ePaperSize.A4;
                    //_genericSheet_PlaCapa.PrinterSettings.FitToPage = true;
                    //_genericSheet_PlaCapa.PrinterSettings.Orientation = eOrientation.Landscape;
                    //_genericSheet_PlaCapa.View.PageBreakView = true;

                    //Actividad actividad = new Actividad();
                    //actividad.iCodExtensionista = fichaTecnica.iCodExtensionista;
                    //DataTable SearTab = capacitacionDL.Listar_PlanCapa_Rpt(actividad);
                    //rowIndexComp = 1;
                    //colcomp = 2;

                    //if (SearTab.Rows.Count > 0)
                    //{
                    //    for (int i = 0; i < SearTab.Rows.Count; i++)
                    //    {
                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "", "#72AEA5");
                    //        //colcomp =+ 7;
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                    //        rowIndexComp++;
                    //        //colcomp =- 7;


                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "PLAN DE CAPACITACIÓN", "#72AEA5");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.Font.Bold = true;
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        rowIndexComp++;
                    //        //colcomp = -7;

                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "", "#72AEA5");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                    //        rowIndexComp++;
                    //        //colcomp = -7;

                    //        // Cabecera de Plan Capacitación
                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "Nombre del SEAR", "#ffffff");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 2, SearTab.Rows[i][0], "#ffffff");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //        //colcomp = 2;
                    //        rowIndexComp++;
                    //        //colcomp--;
                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "Nombre de extensionista", "#ffffff");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 2, SearTab.Rows[i][1], "#ffffff");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //        //colcomp = 2;
                    //        rowIndexComp++;
                    //        //colcomp--;
                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "Agencia Agraria", "#ffffff");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 2, SearTab.Rows[i][3], "#ffffff");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //        //colcomp = 2;
                    //        rowIndexComp++;
                    //        //colcomp--;
                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "Dirección Zonal", "#ffffff");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 2, SearTab.Rows[i][4], "#ffffff");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //        //colcomp = 2;
                    //        rowIndexComp++;
                    //        //colcomp--;
                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, SearTab.Rows[i][8], "#72AEA5");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //        _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 2, SearTab.Rows[i][9], "#72AEA5");
                    //        _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;


                    //        /*****************************************************************************************/
                    //        PlanCapacitacion planCapacitacion = new PlanCapacitacion();
                    //        planCapacitacion.iCodActividad = Convert.ToInt32(SearTab.Rows[i][7]);
                    //        DataTable Modulotab = capacitacionDL.SP_Listar_PlanCapa_Rpt2(planCapacitacion);
                    //        //rowIndexComp = 1;
                    //        //colcomp = 2;
                    //        for (int j = 0; j < Modulotab.Rows.Count; j++)
                    //        {
                    //            rowIndexComp++;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "Modulo o tema", "#ffffff");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 2, Modulotab.Rows[j][0], "#E2EFDA");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;

                    //            rowIndexComp++;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "Objetivo de la sesión", "#ffffff");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 2, Modulotab.Rows[j][1], "#E2EFDA");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;

                    //            //Salto de Linea
                    //            rowIndexComp++;

                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "Meta (productores)", "#ffffff");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 2, Modulotab.Rows[j][2], "#E2EFDA");
                    //            //_genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 3, "Beneficiarios", "#ffffff");
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 4, Modulotab.Rows[j][3], "#E2EFDA");
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 5, "Fecha", "#ffffff");
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 6, Modulotab.Rows[j][4], "#E2EFDA");

                    //            //Salto de Linea
                    //            rowIndexComp++;

                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "Duración Total (horas)", "#ffffff");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 2, Modulotab.Rows[j][7], "#E2EFDA");
                    //            //_genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 3, "Teoria", "#ffffff");
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 4, Modulotab.Rows[j][5], "#E2EFDA");
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 5, "Práctica", "#ffffff");
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 6, Modulotab.Rows[j][6], "#E2EFDA");

                    //            rowIndexComp++;

                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "Lugar de ejecución", "#ffffff");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 2, Modulotab.Rows[j][8], "#E2EFDA");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;

                    //            rowIndexComp++;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "Plan de sesión", "#ffffff");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.Font.Bold = true;
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //            rowIndexComp++;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, "Duración", "#ffffff");
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 1, "Tematica / Pasos", "#ffffff");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 1, rowIndexComp, colcomp + 2].Merge = true;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 3, "Descripción de la Metodología", "#ffffff");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 3, rowIndexComp, colcomp + 5].Merge = true;
                    //            _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 6, "Materiales", "#ffffff");
                    //            _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 6, rowIndexComp, colcomp + 6].Merge = true;

                    //            /*****************************************************************************************/
                    //            PlanSesion planSesion = new PlanSesion();
                    //            planSesion.iCodPlanCap = Convert.ToInt32(Modulotab.Rows[j][9]);
                    //            DataTable SessionModtab = capacitacionDL.SP_Listar_PlanCapa_Rpt3(planSesion);
                    //            for (int k = 0; k < SessionModtab.Rows.Count; k++)
                    //            {
                    //                rowIndexComp++;
                    //                _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp, SessionModtab.Rows[k][0], "#E2EFDA");
                    //                _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 1, SessionModtab.Rows[k][1], "#E2EFDA");
                    //                _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 1, rowIndexComp, colcomp + 2].Merge = true;
                    //                _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 3, SessionModtab.Rows[k][2], "#E2EFDA");
                    //                _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 3, rowIndexComp, colcomp + 5].Merge = true;
                    //                _texto_row(_genericSheet_PlaCapa, rowIndexComp, colcomp + 6, SessionModtab.Rows[k][3], "#E2EFDA");
                    //                _genericSheet_PlaCapa.Cells[rowIndexComp, colcomp + 6, rowIndexComp, colcomp + 6].Merge = true;
                    //            }
                    //        }

                    //        rowIndexComp = 1;
                    //        colcomp = colcomp + 8;

                    //    }
                    //    int totalmetas = 0;
                    //}
                    //#endregion

                    //#region PlanAsistenciaTecnica
                    //NombreReporte = "Plan AT";
                    //var _genericSheet_PlanAT = excelPackage.Workbook.Worksheets.Add(NombreReporte);
                    //_genericSheet_PlanAT.View.ShowGridLines = false;
                    //_genericSheet_PlanAT.View.ZoomScale = 100;
                    //_genericSheet_PlanAT.PrinterSettings.PaperSize = ePaperSize.A4;
                    //_genericSheet_PlanAT.PrinterSettings.FitToPage = true;
                    //_genericSheet_PlanAT.PrinterSettings.Orientation = eOrientation.Landscape;
                    //_genericSheet_PlanAT.View.PageBreakView = true;


                    //SearTab = asistenciaTecDL.SP_Listar_PlanAsistenciaTecnica_Rpt(actividad);
                    //rowIndexComp = 1;
                    //colcomp = 2;

                    //if (SearTab.Rows.Count > 0)
                    //{
                    //    for (int i = 0; i < SearTab.Rows.Count; i++)
                    //    {
                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "", "#72AEA5");
                    //        //colcomp =+ 7;
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                    //        rowIndexComp++;
                    //        //colcomp =- 7;


                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "PLAN DE ASISTENCIA TÉCNICA", "#72AEA5");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.Font.Bold = true;
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //        rowIndexComp++;
                    //        //colcomp = -7;

                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "", "#72AEA5");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                    //        rowIndexComp++;
                    //        //colcomp = -7;

                    //        // Cabecera de Plan Capacitación
                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "Nombre del SEAR", "#ffffff");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 2, SearTab.Rows[i][0], "#ffffff");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //        //colcomp = 2;
                    //        rowIndexComp++;
                    //        //colcomp--;
                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "Nombre de extensionista", "#ffffff");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 2, SearTab.Rows[i][1], "#ffffff");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //        //colcomp = 2;
                    //        rowIndexComp++;
                    //        //colcomp--;
                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "Agencia Agraria", "#ffffff");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 2, SearTab.Rows[i][3], "#ffffff");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //        //colcomp = 2;
                    //        rowIndexComp++;
                    //        //colcomp--;
                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "Dirección Zonal", "#ffffff");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 2, SearTab.Rows[i][4], "#ffffff");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //        //colcomp = 2;
                    //        rowIndexComp++;
                    //        //colcomp--;
                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, SearTab.Rows[i][8], "#72AEA5");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //        _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 2, SearTab.Rows[i][9], "#72AEA5");
                    //        _genericSheet_PlanAT.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;


                    //        /*****************************************************************************************/
                    //        PlanAsistenciaTec planAsistenciaTec = new PlanAsistenciaTec();
                    //        planAsistenciaTec.iCodActividad = Convert.ToInt32(SearTab.Rows[i][7]);
                    //        DataTable Modulotab = asistenciaTecDL.SP_Listar_PlanAsistenciaTecnica_Rpt2(planAsistenciaTec);
                    //        for (int j = 0; j < Modulotab.Rows.Count; j++)
                    //        {
                    //            rowIndexComp++;
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "Objetivo de la sesión", "#ffffff");
                    //            _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 2, Modulotab.Rows[j][1], "#E2EFDA");
                    //            _genericSheet_PlanAT.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;

                    //            //Salto de Linea
                    //            rowIndexComp++;

                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "Meta (productores)", "#ffffff");
                    //            _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 2, Modulotab.Rows[j][2], "#E2EFDA");
                    //            //_genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 3, "Beneficiarios", "#ffffff");
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 4, Modulotab.Rows[j][3], "#E2EFDA");
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 5, "Fecha", "#ffffff");
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 6, Modulotab.Rows[j][4], "#E2EFDA");

                    //            //Salto de Linea
                    //            rowIndexComp++;

                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "Duración Total (horas)", "#ffffff");
                    //            _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 2, Modulotab.Rows[j][7], "#E2EFDA");
                    //            //_genericSheet.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 3, "Teoria", "#ffffff");
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 4, Modulotab.Rows[j][5], "#E2EFDA");
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 5, "Práctica", "#ffffff");
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 6, Modulotab.Rows[j][6], "#E2EFDA");

                    //            rowIndexComp++;

                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "Lugar de ejecución", "#ffffff");
                    //            _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 1].Merge = true;
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 2, Modulotab.Rows[j][8], "#E2EFDA");
                    //            _genericSheet_PlanAT.Cells[rowIndexComp, colcomp + 2, rowIndexComp, colcomp + 6].Merge = true;

                    //            rowIndexComp++;
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "Plan de sesión", "#ffffff");
                    //            _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Merge = true;
                    //            _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.Font.Bold = true;
                    //            _genericSheet_PlanAT.Cells[rowIndexComp, colcomp, rowIndexComp, colcomp + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //            rowIndexComp++;
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, "Duración", "#ffffff");
                    //            //_texto_row(_genericSheet, rowIndexComp, colcomp + 1, "Tematica / Pasos", "#ffffff");
                    //            //_genericSheet.Cells[rowIndexComp, colcomp + 1, rowIndexComp, colcomp + 2].Merge = true;
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 1, "Descripción de la Metodología", "#ffffff");
                    //            _genericSheet_PlanAT.Cells[rowIndexComp, colcomp + 1, rowIndexComp, colcomp + 4].Merge = true;
                    //            _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 5, "Instrumentos", "#ffffff");
                    //            _genericSheet_PlanAT.Cells[rowIndexComp, colcomp + 5, rowIndexComp, colcomp + 6].Merge = true;

                    //            /*****************************************************************************************/
                    //            PlanAsistenciaTecDet planPlanAsistenciaTecDet = new PlanAsistenciaTecDet();
                    //            planPlanAsistenciaTecDet.iCodPlanAsistenciaTec = Convert.ToInt32(Modulotab.Rows[j][9]);
                    //            DataTable SessionModtab = asistenciaTecDL.SP_Listar_PlanAsistenciaTecnica_Rpt3(planPlanAsistenciaTecDet);
                    //            for (int k = 0; k < SessionModtab.Rows.Count; k++)
                    //            {
                    //                rowIndexComp++;
                    //                _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp, SessionModtab.Rows[k][0], "#E2EFDA");
                    //                //_texto_row(_genericSheet, rowIndexComp, colcomp + 1, SessionModtab.Rows[k][1], "#E2EFDA");
                    //                //_genericSheet.Cells[rowIndexComp, colcomp + 1, rowIndexComp, colcomp + 2].Merge = true;
                    //                _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 1, SessionModtab.Rows[k][2], "#E2EFDA");
                    //                _genericSheet_PlanAT.Cells[rowIndexComp, colcomp + 1, rowIndexComp, colcomp + 4].Merge = true;
                    //                _texto_row(_genericSheet_PlanAT, rowIndexComp, colcomp + 5, SessionModtab.Rows[k][3], "#E2EFDA");
                    //                _genericSheet_PlanAT.Cells[rowIndexComp, colcomp + 5, rowIndexComp, colcomp + 6].Merge = true;
                    //            }
                    //        }

                    //        rowIndexComp = 1;
                    //        colcomp = colcomp + 8;

                    //    }
                    //    int totalmetas = 0;
                    //}
                    //#endregion

                    //// Informacion General
                    //#region AnchoColumnas
                    //_genericSheet.Columns[1].Width = 0.38;
                    //_genericSheet.Columns[2].Width = 11.89;
                    //_genericSheet.Columns[3].Width = 11.67;
                    //_genericSheet.Columns[4].Width = 27.56;
                    //_genericSheet.Columns[5].Width = 24.22;
                    //_genericSheet.Columns[6].Width = 13.22;
                    //_genericSheet.Columns[7].Width = 10.78;
                    //_genericSheet.Columns[8].Width = 13.33;
                    //_genericSheet.Columns[9].Width = 16.56;
                    //_genericSheet.Columns[10].Width = 10.56;
                    //#endregion
                    //// Identificacion
                    //#region AnchoColumnas
                    //_genericSheet_.Columns[1].Width = 1.11;
                    //_genericSheet_.Columns[2].Width = 13.56;
                    //_genericSheet_.Columns[3].Width = 14.56;
                    //_genericSheet_.Columns[4].Width = 13.22;
                    //_genericSheet_.Columns[5].Width = 25.33;
                    //_genericSheet_.Columns[6].Width = 11.33;
                    //_genericSheet_.Columns[7].Width = 14.56;
                    //_genericSheet_.Columns[8].Width = 21.22;
                    //_genericSheet_.Columns[9].Width = 16.22;
                    //#endregion

                    #endregion



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