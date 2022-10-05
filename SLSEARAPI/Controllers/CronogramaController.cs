using OfficeOpenXml;
using OfficeOpenXml.Style;
using SLSEARAPI.DataLayer;
using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using System.Windows.Media;

namespace SLSEARAPI.Controllers
{
    public class CronogramaController : ApiController
    {
        CronogramaDL cronogramaDL;
        public CronogramaController()
        {
            cronogramaDL = new CronogramaDL();
        }

        [HttpPost]
        [ActionName("ListarComponentes")]
        public List<ComponenteCronograma> ListarComponentes(Identificacion indicadores)
        {
            try
            {
                return cronogramaDL.ListarComponentes(indicadores);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListaActividades")]
        public List<Actividad> ListaActividades(Identificacion identificacion)
        {
            try
            {
                return cronogramaDL.ListaActividades(identificacion);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("InsertarCronograma")]
        public Cronograma InsertarCronograma(Cronograma cronograma)
        {
            try
            {
                return cronogramaDL.InsertarCronograma(cronograma);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ListarCronograma")]
        public List<Cronograma> ListarCronograma(Cronograma cronograma)
        {
            try
            {
                return cronogramaDL.ListarCronograma(cronograma);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ExportarCronograma")]
        public HttpResponseMessage ExportarCronograma(Cronograma cronograma)
        {
            try
            {
                String NombreReporte = "cronograma";

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
                    _cabecera.Add("Actividad");
                    _cabecera.Add("Descripcion");                    
                    _cabecera.Add("Unidad Medida");
                    _cabecera.Add("Meta");

                    int finish = pintarcabeceras(_cabecera, _genericSheet, "Cronograma Detallado");

                    _genericSheet.Cells["A3:D3"].Value = "3.1 CRONOGRAMA DE EJECUCIÓN";
                    _genericSheet.Cells["A3:D3"].Merge = true;
                    _genericSheet.Cells["A3:D3"].Style.Locked = true;
                    System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#72aea5");
                    _genericSheet.Cells["A3:D3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _genericSheet.Cells["A3:D3"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    _genericSheet.Column(1).Style.Locked = true;
                    _genericSheet.Workbook.Protection.LockWindows = true;
                    _genericSheet.Workbook.Protection.LockStructure = true;

                    //pintar componente 1

                    List<Componente> componentes = cronogramaDL.ListarComponentes(cronograma);
                    int rowIndexComp = 7;
                    int colcomp = 1;
                                        
                    _texto_row(_genericSheet, rowIndexComp, colcomp++,1, "#fff2cc");
                    _texto_row(_genericSheet, rowIndexComp, colcomp++, "Componente 1", "#fff2cc");
                    _texto_row(_genericSheet, rowIndexComp, colcomp++, componentes[0].vDescComponente, "#fff2cc");
                    _texto_row(_genericSheet, rowIndexComp, colcomp++, componentes[0].vUnidadMedida, "#fff2cc");
                    _texto_row(_genericSheet, rowIndexComp, colcomp++,int.Parse(componentes[0].vMeta), "#fff2cc");

                    // Suma de Totales de Meta de Compromisos
                    int totalmetas = 0;

                    foreach (Componente item in componentes)
                    {
                        totalmetas=totalmetas+int.Parse(item.vMeta);
                    }


                    List<Cronograma> lista = cronogramaDL.ListarConogramaFecha(cronograma);
                    Actividad actividad = new Actividad();
                    actividad.iCodExtensionista = cronograma.iCodExtensionista;
                    actividad.nTipoActividad = 1;
                    List<Actividad> listaactividades = cronogramaDL.ListarActividades(actividad);
                    int rowIndex = 8;
                    int Nro = 2;

                    //// pintar las actividades

                    foreach (var d in listaactividades)
                    {
                        int col89 = 1;
                        _texto_row(_genericSheet, rowIndex, col89++,decimal.Parse("1."+Nro.ToString()));
                        _texto_row(_genericSheet, rowIndex, col89++, d.vActividad);
                        _texto_row(_genericSheet, rowIndex, col89++, d.vDescripcion);
                        _texto_row(_genericSheet, rowIndex, col89++, d.vUnidadMedida);
                        _texto_row(_genericSheet, rowIndex, col89++,int.Parse(d.vMeta));

                        Nro = Nro + 1;
                        rowIndex++;
                    }
                                    
                    ////pintar las fechas 

                    int rowIndex5 = 5;
                    int rowFila = 6;
                    
                    int Nro1 = 1;                    
                    int col2 = 6;
                    int col1 = 6;

                    IEnumerable<DateTime> listafechas = lista.Select(a => a.dfechacronograma).Distinct();
                                        
                    foreach (var item in listafechas)
                    {                    
                        _texto_row_fecha(_genericSheet, rowIndex5, col1++,DateTime.Parse(item.ToString("dd/MM/yyyy")),"d-mmm");
                        _texto_row_fecha(_genericSheet, rowFila, col2++, Nro1);                        
                        Nro1++;
                    }

                  
                //// pintar cotenido
                ///
                cronograma.nTipoActividad = 1;
                lista = cronogramaDL.ListarConogramaFechaTipo(cronograma);
                int coldata = 6;

                int rownext = 0;

                List<int> Cantidades = new List<int>();

                foreach (var itemfecha in listafechas)
                {
                    int rowdata = 8;
                    foreach (var item in lista)
                    {
                        if (itemfecha.ToString("dd/MM/yyyy") == item.dfechacronograma.ToString("dd/MM/yyyy") && item.iCodComponente == 1)
                        {
                                Cantidades.Add(item.iCantidad);
                                _texto_row(_genericSheet, rowdata, coldata, item.iCantidad, "#e2efda");  
                        }
                        rowdata++;
                    }
                    coldata++;
                    rownext = rowdata;
                }

                    // Pintar totales del componente 1
                    coldata = 6;
                    foreach (int item in Cantidades)
                    {
                        _texto_row(_genericSheet, rowIndexComp, coldata++, item, "#fff2cc");
                    }

                    // pintar componentes 2

                  int rowcomponente2 = rowIndex;

                  _texto_row(_genericSheet, rowcomponente2, 1, 2, "#fff2cc");
                  _texto_row(_genericSheet, rowcomponente2, 2, "Componente 2", "#fff2cc");
                  _texto_row(_genericSheet, rowcomponente2, 3, componentes[1].vDescComponente, "#fff2cc");
                  _texto_row(_genericSheet, rowcomponente2, 4, componentes[1].vUnidadMedida, "#fff2cc");
                  _texto_row(_genericSheet, rowcomponente2, 5,int.Parse(componentes[1].vMeta), "#fff2cc");

                    int rowcomponente21 = rowcomponente2;

                    //// actividades 2



                    Actividad actividad1 = new Actividad();
                 actividad1.iCodExtensionista = cronograma.iCodExtensionista;
                 actividad1.nTipoActividad = 2;

                 List<Actividad> listaactividades2 = cronogramaDL.ListarActividades(actividad1);

                 Nro = 1;
                 rowIndex = rowcomponente2+1;
                                     
                 foreach (var d in listaactividades2)
                 {
                     int col89 = 1;
                     _texto_row(_genericSheet, rowIndex, col89++, decimal.Parse("2." + Nro.ToString()));
                     _texto_row(_genericSheet, rowIndex, col89++, d.vActividad);
                     _texto_row(_genericSheet, rowIndex, col89++, d.vDescripcion);
                     _texto_row(_genericSheet, rowIndex, col89++, d.vUnidadMedida);
                     _texto_row(_genericSheet, rowIndex, col89++, int.Parse(d.vMeta));

                     Nro = Nro + 1;
                     rowIndex++;
                    }

                    // Mostrar el total de metas de los componentes

                    _texto_row(_genericSheet, rowIndex, 1, "");
                    _texto_row(_genericSheet, rowIndex, 2, "");
                    _texto_row(_genericSheet, rowIndex, 3, "");
                    _texto_row(_genericSheet, rowIndex, 4, "");
                    _texto_row(_genericSheet, rowIndex, 5, totalmetas);

                    cronograma.nTipoActividad = 2;
                    //List<Cronograma> lista1 = cronogramaDL.ListarConogramaFechaTipo(cronograma);
                    

                    lista = cronogramaDL.ListarConogramaFechaTipo(cronograma);
                    IEnumerable<DateTime> listafechas2 = lista.Select(a => a.dfechacronograma).Distinct();
                    //listafechas = lista.Select(a => a.dFecha).Distinct();

                    listafechas.ToList().AddRange(listafechas2);
                    //lista.AddRange(lista1);
                    ////coldata = 6;
                    //rownext = rowcomponente2+1;
                    //int cantidad_actividades = listaactividades2.Count();
                    int rowdata1 = rowcomponente2 + 1;
                    coldata = 6;

                    int totalcomponente1 = Cantidades.Count;

                    foreach (Actividad itemact in listaactividades2)
                    {                        
                        foreach (var itemfecha in listafechas)
                        {
                            foreach (var item in lista)
                            {
                                if (itemfecha.ToString("dd/MM/yyyy") == item.dfechacronograma.ToString("dd/MM/yyyy") && item.iCodActividad == itemact.iCodActividad)
                                {
                                    Cantidades.Add(item.iCantidad);
                                    _texto_row(_genericSheet, rowdata1, coldata, item.iCantidad, "#e2efda");
                                }                                    
                            }
                            coldata++;
                        }
                        rowdata1++;
                        coldata = 6;
                    }
                    //rowIndex++;
                    int colmetas = 6;

                    //Mostrar totales
                    List<int> CantidadesComp2 = new List<int>();

                    //rowcomponente21
                    int contador = 1;

                    foreach (int item in Cantidades)
                    {
                        if(contador>totalcomponente1)
                        {
                            _texto_row(_genericSheet, rowcomponente21, colmetas++, item);
                        }
                        else
                        {
                            _texto_row(_genericSheet, rowcomponente21, colmetas++, "-");
                        }                        
                        contador++;
                    }

                    colmetas = 6;

                    foreach (int item in Cantidades)
                    {                        
                        _texto_row(_genericSheet, rowIndex, colmetas++, item);
                    }

                    //foreach (var itemfecha in listafechas)
                    //{
                       
                    //    //foreach (var itemact in listaactividades2)
                    //    //{                          
                    //        foreach (var item in lista)
                    //        {
                    //            if (itemfecha == item.dFecha)
                    //            {
                    //                //_texto_row(_genericSheet, rowdata1, coldata, "fila:"+ rowdata1+"-columna:"+coldata, "#e2efda");
                    //                _texto_row(_genericSheet, rowdata1, coldata, item.iCantidad, "#e2efda");
                    //            }
                    //            rowdata1++;
                    //            coldata++;
                    //        }
                    //        rowdata1 = rowcomponente2 + 1;
                    //        coldata = 6;
                    //    //}                       
                    //}

                    // funcionaba

                    //foreach (var itemfecha in listafechas)
                    //{
                    //    int rowdata = rowcomponente2 + 1;
                    //    foreach (var itemact in listaactividades2)
                    //    {
                    //        coldata = 6;
                    //        foreach (var item in lista)
                    //        {
                    //            if (itemfecha == item.dFecha && item.iCodActividad == itemact.iCodActividad)
                    //            {
                    //                _texto_row(_genericSheet, rowdata, coldata, item.iCantidad, "#e2efda");
                    //            }
                    //            coldata++;
                    //        }
                    //    }
                    //    rowdata++;
                    //}

                    /*
             */

                    _genericSheet.Column(2).Width = 50d;
                    _genericSheet.Column(3).Width = 70d;
                    _genericSheet.Column(4).Width = 20d;

                    _genericSheet.Row(5).Height = 45d;

                    //_genericSheet.Protection.SetPassword("123456");

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
            _sheet.Cells[rowIndex, col].Style.Font.Name = fontName;
            _sheet.Cells[rowIndex, col].Style.Font.Size = 12;
            _sheet.Cells[rowIndex, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            _sheet.Cells[rowIndex, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[rowIndex, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[rowIndex, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            if(Color!="")
            {
                System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml(Color);
                _sheet.Cells[rowIndex, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                _sheet.Cells[rowIndex, col].Style.Fill.BackgroundColor.SetColor(colFromHex);
            }            
        }

        internal void _texto_row_fecha(ExcelWorksheet _sheet, int rowIndex, int col, object _text,string formato="", string fontName = "Calibri")
        {
            _sheet.Cells[rowIndex, col].Value = _text;
            _sheet.Cells[rowIndex, col].Merge = true;
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
                _sheet.Cells[rowIndex, col].Style.TextRotation = 90;
            }            
        }
        protected int pintarcabeceras(List<string> _cabecera, ExcelWorksheet _Sheet, string header)
        {
            int finish = -1;
            foreach (var x in _cabecera)
            {
                finish++;
                string letra = convertNumberToLetter(finish);
                setHeader1(_Sheet, letra + "6:" + letra + "6", x, false);
            }
            _texto_sin_borde_Titulo1(_Sheet, convertNumberToLetter(0) + "1:" + convertNumberToLetter(finish) + "1", header, System.Drawing.Color.White, System.Drawing.Color.Black, "Calibri");
            return finish;
        }
        internal void _texto_sin_borde_Titulo(ExcelWorksheet _sheet, String _range, String _text, System.Drawing.Color _Backcolor, System.Drawing.Color _fontColor, string fontName = "Calibri")
        {
            _sheet.Cells[_range].Value = _text;
            _sheet.Cells[_range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            _sheet.Cells[_range].Style.Fill.BackgroundColor.SetColor(_Backcolor);
            _sheet.Cells[_range].Style.Font.Color.SetColor(_fontColor);
            _sheet.Cells[_range].Style.Font.Bold = true;
            _sheet.Cells[_range].Merge = true;
            _sheet.Cells[_range].Style.WrapText = true;
            _sheet.Cells[_range].Style.Font.Size = 14;
            _sheet.Cells[_range].Style.Font.Name = fontName;
            _sheet.Cells[_range].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            _sheet.Cells[_range].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
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
            _sheet.Cells[_range].Style.WrapText = true;
            _sheet.Cells[_range].Style.Font.Size = 18;
            _sheet.Cells[_range].Style.Font.Name = fontName;
            _sheet.Cells[_range].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            _sheet.Cells[_range].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
        public void setHeader(ExcelWorksheet _sheet, string celda, string texto, Boolean mergue)
        {
            _sheet.Cells[celda].Value = texto;
            _sheet.Cells[celda].Style.Font.Name = "Calibri";
            _sheet.Cells[celda].Style.Font.Size = 12;
            _sheet.Cells[celda].Style.Font.Bold = true;
            _sheet.Cells[celda].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            _sheet.Cells[celda].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            _sheet.Cells[celda].Merge = mergue;
            _sheet.Cells[celda].Style.WrapText = true;
            _sheet.Cells[celda].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[celda].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[celda].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[celda].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            _sheet.Cells[celda].Style.Fill.PatternType = ExcelFillStyle.Solid;            
            _sheet.Cells[celda].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(30, 144, 255));
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
            _sheet.Cells[celda].Style.WrapText = true;
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
    }
}
