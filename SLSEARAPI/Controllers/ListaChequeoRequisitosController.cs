using iTextSharp.text;
using iTextSharp.text.pdf;
using SLSEARAPI.DataLayer;
using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Web;
using System.Web.Http;

namespace SLSEARAPI.Controllers
{
    public class ListaChequeoRequisitosController : ApiController
    {
        ListaChequeoRequisitosDL listaChequeoRequisitosDL;

        iTextSharp.text.Font ARIAL13 = new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        iTextSharp.text.Font ARIAL13bLACK = new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 13, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        iTextSharp.text.Font ARIAL8 = new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
        iTextSharp.text.Font ARIAL8bLACK = new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        iTextSharp.text.Font ARIAL8white = new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8, iTextSharp.text.Font.BOLD, BaseColor.WHITE);


        private Exception ex;
        public ListaChequeoRequisitosController()
        {

            listaChequeoRequisitosDL = new ListaChequeoRequisitosDL();
        }

        [HttpPost]
        [ActionName("InsertarListaChequeoRequisitos")]

        public ListaChequeoRequisitos InsertarListaChequeoRequisitos(ListaChequeoRequisitos ListaChequeRequisitos)
        {
            try
            {
                return listaChequeoRequisitosDL.InsertarListaChequeoRequisitos(ListaChequeRequisitos);                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [HttpPost]
        [ActionName("ActualizarDardeBajaListaChequeoRequisitos")]
        public ListaChequeoRequisitos ActualizarDardeBajaListaChequeoRequisitos(ListaChequeoRequisitos entidad)
        {
            try
            {
                return listaChequeoRequisitosDL.ActualizarDardeBajaListaChequeoRequisitos(entidad);                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        [HttpPost]
        [ActionName("ListarChequeoRequisitos")]
        public List<ListaChequeoRequisitos> ListarChequeoRequisitos(Extensionista extensionista)
        {
            try
            {
                return listaChequeoRequisitosDL.ListarChequeoRequisitos(extensionista);                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public HttpResponseMessage GenerarPdfRequisitos(Extensionista extensionista)
        {   
            List<ListaChequeoRequisitos> listarequisitos= listaChequeoRequisitosDL.ListarChequeoRequisitos(extensionista);

            using (MemoryStream memoryStream = new MemoryStream())
            {
                try
                {
                    Document document = new Document(PageSize.A4, 40, 40, 100, 50);
                    PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);
                    //StringBuilder str = new StringBuilder();
                    //str.Append("PP 0068   : REDUCCIÓN DE LA VULNERABILIDAD Y ATENCIÓN DE EMERGENCIAS POR DESASTRES\n");
                    //str.Append("PRODUCTO  : 300735 DESARROLLO DE MEDIDAS DE INTERVENCIÓN PARA LA PROTECCIÓN FÍSICA FRENTE A PELIGROS\n");
                    //str.Append("ACTIVIDAD : 5005865 DESARROLLO DE TÉCNICAS AGROPECUARIAS ANTE PELIGROS HIDROMETEOROLÓGICOS\n");
                    //str.Append("TAREA     : MÓDULO DE PROTECCIÓN DE GANADO ( COBERTIZO )");
                    //writer.PageEvent = new ITextEventsRetencion(HttpContext.Current.Server.MapPath("~/Image") + "/agrorural.jpg", str.ToString());
                    //document.AddTitle("REGISTRO DE AVANCE DE COBERTIZOS");
                    //document.AddCreator("OA/UTI - AGRO RURAL");
                    document.Open();
                    
                    string path = HttpContext.Current.Server.MapPath("~/Content/Images");
                    string imageURL = path + "/logoagrorural2.png";
                    iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imageURL);
                    //Resize image depend upon your need
                    jpg.ScaleToFit(150, 50);
                    //Give space before image
                    //jpg.SpacingBefore = 10f;
                    //Give some space after the image
//                    jpg.SpacingAfter = 10f;
                    jpg.Alignment = Element.ALIGN_RIGHT;
                    jpg.SetAbsolutePosition(420, 770);

                    document.Add(jpg);

                    string imageURL1 = path + "/logomdar.jpg";
                    iTextSharp.text.Image jpg1 = iTextSharp.text.Image.GetInstance(imageURL1);
                    //Resize image depend upon your need
                    jpg1.ScaleToFit(150, 50);
                    //Give space before image
                    //jpg.SpacingBefore = 10f;
                    //Give some space after the image
                    //                    jpg.SpacingAfter = 10f;
                    jpg1.Alignment = Element.ALIGN_RIGHT;
                    jpg1.SetAbsolutePosition(20, 770);
                    document.Add(jpg1);

                    PdfPTable tableTitutlo = new PdfPTable(1);
                    float[] widths = new float[] { 100f };
                    tableTitutlo.SetWidths(widths);
                    PdfPCell pdfCell2 = new PdfPCell(new Phrase(new Chunk("Anexo N° 01. Lista de chequeo de requisitos", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell2.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell2.VerticalAlignment = Element.ALIGN_CENTER;
                    pdfCell2.Border = 0;
                    tableTitutlo.AddCell(pdfCell2);
                    document.Add(tableTitutlo);

                    document.Add(new Paragraph("\n"));

                    PdfPTable pdfCuerpo1 = new PdfPTable(1);
                    float[] widths1 = new float[] { 150f };
                    pdfCuerpo1.SetWidths(widths1);
                    pdfCuerpo1.WidthPercentage = 100f;
                    pdfCuerpo1.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                    PdfPCell pdfCell21 = new PdfPCell(new Phrase(new Chunk("La ALIANZA ESTRATÉGICA está conformada por el proveedor del SEAR (extensionista) y los productores organizados (beneficiarios)", ARIAL8)));

                    pdfCell21.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell21.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell21.Border = 0;

                    pdfCuerpo1.AddCell(pdfCell21);

                    PdfPCell pdfCell22 = new PdfPCell(new Phrase(new Chunk("La ACREDITACIÓN, es un compromiso asumido entre el proveedor del SEAR (extensionista) y los productores beneficiarios a fin de cumplir con la modalidad de postulación de las bases.", ARIAL8)));
                    pdfCell22.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell22.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell22.Border = 0;
                    pdfCuerpo1.AddCell(pdfCell22);

                    ExtensionistaDL extensionistaDL = new ExtensionistaDL();

                    Extensionista extensionistadatos = extensionistaDL.ListarExtensionistaPorCodigo(extensionista);

                    PdfPCell pdfCell23 = new PdfPCell(new Phrase(new Chunk("Nombre de la Propuesta : "+ extensionistadatos.vNombrePropuesta, ARIAL8)));
                    pdfCell23.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell23.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell23.Border = 0;
                    pdfCuerpo1.AddCell(pdfCell23);

                    document.Add(pdfCuerpo1);

                    document.Add(new Paragraph("\n"));

                    // requisitos 

                    PdfPTable pdfCuerpo33 = new PdfPTable(3);
                    float[] widths123 = new float[] { 150f, 25F, 25f };

                    pdfCuerpo33.SetWidths(widths123);
                    pdfCuerpo33.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                    //pdfCuerpo123.WidthPercentage = 100f;
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("REQUISITOS",ARIAL8,1,1,0,false));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("SI CUMPLE", ARIAL8,1,1,0,false));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("NO CUMPLE", ARIAL8,1,1,0,false));
                                    
                    Boolean respuestacumpleA= BuscarSeleccionadoRequisito(listarequisitos, 1);

                    pdfCuerpo33.AddCell(GenerateTableUnaFila("A. Que la propuesta presentada no es copia textual de otras propuestas evaluadas, en ejecución o culminados de otras fuentes de financiamiento", ARIAL8,0,2,0,false));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8,0,1,1,respuestacumpleA));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8,0,1,1,!respuestacumpleA));

                    Boolean respuestacumpleB = BuscarSeleccionadoRequisito(listarequisitos, 2);

                    pdfCuerpo33.AddCell(GenerateTableUnaFila("B. Que el proveedor del SEAR no está observado por otra fuente de financiamiento a la que se tenga acceso por un mal desempeño o incumpliendo contractual en un proyecto culminado o en ejecución.", ARIAL8, 0, 2,0,false));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8, 0, 1,1, respuestacumpleB));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8, 0, 1,1,!respuestacumpleB));

                    Boolean respuestacumpleC = BuscarSeleccionadoRequisito(listarequisitos, 3);

                    pdfCuerpo33.AddCell(GenerateTableUnaFila("C. Que el proveedor del SEAR no se encuentre impedido de contratar con el Estado.", ARIAL8, 0, 2,0,false));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8, 0, 1,1, respuestacumpleC));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8, 0, 1,1, !respuestacumpleC));

                    Boolean respuestacumpleD = BuscarSeleccionadoRequisito(listarequisitos, 4);

                    pdfCuerpo33.AddCell(GenerateTableUnaFila("D. Que el proveedor del SEAR no presenta deudas coactivas con el Estado reportadas por la SUNAT", ARIAL8, 0, 2,0,false));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8, 0, 1,1, respuestacumpleD));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8, 0, 1,1, !respuestacumpleD));

                    Boolean respuestacumpleE = BuscarSeleccionadoRequisito(listarequisitos, 5);

                    pdfCuerpo33.AddCell(GenerateTableUnaFila("E.Que el proveedor del SEAR cumple con las condiciones establecidas en el Ítems 1.10.1 condiciones del proveedor, de las presentes bases.", ARIAL8, 0, 2,0,false));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8, 0, 1,1, respuestacumpleE));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8, 0, 1,1, !respuestacumpleE));

                    Boolean respuestacumpleF = BuscarSeleccionadoRequisito(listarequisitos, 6);

                    pdfCuerpo33.AddCell(GenerateTableUnaFila("F. Que los beneficiarios cumplen con las condiciones establecidas en el Ítems 1.10.2 condiciones del proveedor, de las presentes bases.", ARIAL8, 0, 2,0,false));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8, 0, 1,1, respuestacumpleF));
                    pdfCuerpo33.AddCell(GenerateTableUnaFila("", ARIAL8, 0, 1,1, !respuestacumpleF));

                    document.Add(pdfCuerpo33);

                    document.Add(new Paragraph("\n"));


                    PdfPTable pdfCuerpo44 = new PdfPTable(1);
                    float[] widths4 = new float[] { 150f };
                    pdfCuerpo44.SetWidths(widths4);
                    //pdfCuerpo44.WidthPercentage = 100f;
                    pdfCuerpo44.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                    PdfPCell pdfCell41 = new PdfPCell(new Phrase(new Chunk("TENER EN CUENTA:", ARIAL8bLACK)));

                    pdfCell41.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell41.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell41.Border = 0;
                    pdfCuerpo44.AddCell(pdfCell41);
                    
                    PdfPCell pdfCell42 = new PdfPCell(new Phrase(new Chunk("Si no cumple con algunos de requisitos, absténgase de participar. El presente documento forma parte integral de las bases de la presente convocatoria, y su contenido no puede ser modificado parcial o totalmente.", ARIAL8)));

                    pdfCell42.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell42.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell42.Border = 0;
                    pdfCuerpo44.AddCell(pdfCell42);

                    PdfPCell pdfCell43 = new PdfPCell(new Phrase(new Chunk("En caso de descubrirse que la información entregada es falsa o no estuviera acreditada, el proveedor del SEAR (extensionista) y la propuesta quedan descalificados en forma automática.", ARIAL8)));

                    pdfCell43.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell43.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell43.Border = 0;
                    pdfCuerpo44.AddCell(pdfCell43);

                    PdfPCell pdfCell44 = new PdfPCell(new Phrase(new Chunk("AGRO RURAL tiene la facultad de realizar la fiscalización o validación, en cualquier etapa del concurso, de la información declarada en el presente formato, según Ley de Procedimiento Administrativo General.", ARIAL8)));

                    pdfCell44.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell44.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell44.Border = 0;
                    pdfCuerpo44.AddCell(pdfCell44);

                    PdfPCell pdfCell64 = new PdfPCell(new Phrase(new Chunk("-", ARIAL8white)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell64.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell64.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell64.Border = 0;
                    pdfCuerpo44.AddCell(pdfCell64);

                    PdfPCell pdfCell65 = new PdfPCell(new Phrase(new Chunk("-", ARIAL8white)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell65.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell65.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell65.Border = 0;
                    pdfCuerpo44.AddCell(pdfCell65);

                    PdfPCell pdfCell66 = new PdfPCell(new Phrase(new Chunk("-", ARIAL8white)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell66.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell66.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell66.Border = 0;
                    pdfCuerpo44.AddCell(pdfCell66);

                    PdfPCell pdfCell61 = new PdfPCell(new Phrase(new Chunk("________________________________________", ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell61.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell61.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell61.Border = 0;
                    pdfCuerpo44.AddCell(pdfCell61);

                    PdfPCell pdfCell62 = new PdfPCell(new Phrase(new Chunk("Nombre y Apellido: " + extensionistadatos.vNombres + " " + extensionistadatos.vApemat + " " + extensionistadatos.vApepat, ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell62.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell62.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell62.Border = 0;
                    pdfCuerpo44.AddCell(pdfCell62);

                    PdfPCell pdfCell63 = new PdfPCell(new Phrase(new Chunk("Proveedor del SEAR – Extensionista", ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell63.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell63.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell63.Border = 0;
                    pdfCuerpo44.AddCell(pdfCell63);

                    document.Add(pdfCuerpo44);                                      

                    document.Add(new Paragraph("\n"));

              
                    document.Close();

                    byte[] buffer = memoryStream.ToArray();
                    var contentLength = buffer.Length;
                    var result = Request.CreateResponse(HttpStatusCode.OK);
                    result.Content = new StreamContent(new MemoryStream(buffer));
                    result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = Guid.NewGuid().ToString() + "_" + DateTime.Now.ToShortDateString() + ".pdf"
                    };
                    return result;
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
        public HttpResponseMessage GenerarPdfActaCompromiso(Extensionista extensionista)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                try
                {
                    Document document = new Document(PageSize.A4, 40, 40, 100, 50);
                    PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);

                    //writer.PageEvent = new ITextEventsRetencion(HttpContext.Current.Server.MapPath("~/Image") + "/agrorural.jpg", str.ToString());                                        
                    document.Open();

                    string path = HttpContext.Current.Server.MapPath("~/Content/Images");
                    string imageURL = path + "/logoagrorural2.png";
                    iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imageURL);
                    //Resize image depend upon your need
                    jpg.ScaleToFit(150, 50);
                    //Give space before image
                    //jpg.SpacingBefore = 10f;
                    //Give some space after the image
                    //                    jpg.SpacingAfter = 10f;
                    jpg.Alignment = Element.ALIGN_RIGHT;
                    jpg.SetAbsolutePosition(420, 770);

                    document.Add(jpg);

                    string imageURL1 = path + "/logomdar.jpg";
                    iTextSharp.text.Image jpg1 = iTextSharp.text.Image.GetInstance(imageURL1);
                    //Resize image depend upon your need
                    jpg1.ScaleToFit(150, 50);
                    //Give space before image
                    //jpg.SpacingBefore = 10f;
                    //Give some space after the image
                    //                    jpg.SpacingAfter = 10f;
                    jpg1.Alignment = Element.ALIGN_RIGHT;
                    jpg1.SetAbsolutePosition(20, 770);
                    document.Add(jpg1);

                    PdfPTable tableTitutlo = new PdfPTable(1);
                    float[] widths = new float[] { 100f };
                    tableTitutlo.SetWidths(widths);
                    PdfPCell pdfCell2 = new PdfPCell(new Phrase(new Chunk("Anexo N° 03. Acta de Compromiso de Asociación", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell2.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell2.VerticalAlignment = Element.ALIGN_CENTER;
                    pdfCell2.Border = 0;
                    tableTitutlo.AddCell(pdfCell2);
                    document.Add(tableTitutlo);

                    document.Add(new Paragraph("\n"));

                    ExtensionistaDL extensionistaDL = new ExtensionistaDL();

                    Extensionista extensionistadatos = extensionistaDL.ListarExtensionistaPorCodigo(extensionista);

                    PdfPTable pdfCuerpo1 = new PdfPTable(1);
                    float[] widths1 = new float[] { 150f };
                    pdfCuerpo1.SetWidths(widths1);
                    pdfCuerpo1.WidthPercentage = 100f;
                    pdfCuerpo1.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                    PdfPCell pdfCell21 = new PdfPCell(new Phrase(new Chunk("En la ciudad de "+ extensionistadatos .vNomDistrito+ ", Provincia de "+extensionistadatos.vNomProvincia+", y Departamento de "+extensionistadatos.vNomDepartamento+"; siendo las ………….. horas del día ………… del mes de ……………. del año 2022, los señores productores:", ARIAL8)));

                    pdfCell21.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell21.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell21.Border = 0;

                    pdfCuerpo1.AddCell(pdfCell21);

                    document.Add(pdfCuerpo1);

                    document.Add(new Paragraph("\n"));

                    ActaAlianzaEstrategicaDL actaAlianzaEstrategicaDL = new ActaAlianzaEstrategicaDL();

                    // productores

                    PdfPTable pdfCuerpo33 = new PdfPTable(8);
                    float[] widths123 = new float[] { 10f, 100f, 25F, 30f, 20f, 30f, 55f, 25f };

                    pdfCuerpo33.SetWidths(widths123);
                    pdfCuerpo33.WidthPercentage = 100f;
                    //pdfCuerpo33.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                    PdfPCell pdfCell31 = new PdfPCell(new Phrase(new Chunk("N°", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell31.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell31.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell31.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell31);

                    PdfPCell pdfCell32 = new PdfPCell(new Phrase(new Chunk("Apellidos y Nombres", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell32.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell32.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell32.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell32);

                    PdfPCell pdfCell33 = new PdfPCell(new Phrase(new Chunk("DNI", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell33.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell33.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell33.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell33);

                    PdfPCell pdfCell34 = new PdfPCell(new Phrase(new Chunk("Celular", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell34.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell34.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell34.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell34);

                    PdfPCell pdfCell35 = new PdfPCell(new Phrase(new Chunk("Edad", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell35.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell35.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell35.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell35);

                    PdfPCell pdfCell36 = new PdfPCell(new Phrase(new Chunk("Sexo", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell36.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell36.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell36.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell36);

                    PdfPCell pdfCell37 = new PdfPCell(new Phrase(new Chunk("Recibo Capacitacion y/o asistencia técnica", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell37.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell37.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell37.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell37);

                    PdfPCell pdfCell38 = new PdfPCell(new Phrase(new Chunk("Firma", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell38.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell38.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell38.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell38);
                                        
                    List<Productor> lista = new List<Productor>();
                    Productor productor = new Productor();

                    productor.iCodExtensionista = extensionista.iCodExtensionista;
                    productor.piPageSize = 40;
                    productor.piCurrentPage = 1;
                    productor.pvSortColumn = "iCodProductor";
                    productor.pvSortOrder = "asc";
                    productor.iPerteneceOrganizacion = 0;

                    lista = actaAlianzaEstrategicaDL.ListarProductor(productor);
                    int i = 0;
                    foreach (Productor item in lista)
                    {
                        PdfPCell pdfCell39 = new PdfPCell(new Phrase(new Chunk((i + 1).ToString(), ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell39.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell39.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell39);

                        PdfPCell pdfCell40 = new PdfPCell(new Phrase(new Chunk(item.vApellidosNombres, ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell40.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell40.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell40);

                        PdfPCell pdfCell41 = new PdfPCell(new Phrase(new Chunk(item.vDni, ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell41.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell41.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell41);

                        PdfPCell pdfCell45 = new PdfPCell(new Phrase(new Chunk(item.vCelular, ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell45.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell45.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell45);


                        PdfPCell pdfCell42 = new PdfPCell(new Phrase(new Chunk(item.iEdad.ToString(), ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell42.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell42.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell42);

                        PdfPCell pdfCell43 = new PdfPCell(new Phrase(new Chunk(item.iSexo == 1 ? "Masculino" : "Femenino", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell43.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell43.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell43);

                        PdfPCell pdfCell44 = new PdfPCell(new Phrase(new Chunk(item.iRecibioCapacitacion==1 ? "SI":"NO", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell44.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell44.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell44);

                        PdfPCell pdfCell46 = new PdfPCell(new Phrase(new Chunk("", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell46.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell46.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell46);
                        i++;
                    }

                    document.Add(pdfCuerpo33);

                    document.Add(new Paragraph("\n"));

                    PdfPTable pdfCuerpo55 = new PdfPTable(1);

                    float[] widths5 = new float[] { 100f };

                    pdfCuerpo55.SetWidths(widths5);
                    pdfCuerpo55.WidthPercentage = 100f;

                    //PdfPCell pdfCell51 = new PdfPCell(new Phrase(new Chunk("En calidad de “Beneficiarios” por mutuo acuerdo deciden participar en el Servicios de Extensión Agraria Rural SEAR 2022.", ARIAL8)));
                    ////pdfCell2.BackgroundColor = BaseColor.BLACK;
                    //pdfCell51.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    //pdfCell51.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    //pdfCell31.Border = 0;
                    //pdfCuerpo55.AddCell(pdfCell51);

                    //PdfPCell pdfCell52 = new PdfPCell(new Phrase(new Chunk("Así mismo, el Sr. ……………………………………………………………………………. Identificado con número de DNI Nº …………………. y RUC Nº …………………………..…….decide participar en calidad de “Proveedor del SEAR – Extensionista”.", ARIAL8)));
                    ////pdfCell2.BackgroundColor = BaseColor.BLACK;
                    //pdfCell52.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    //pdfCell52.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    //pdfCell31.Border = 0;
                    //pdfCuerpo55.AddCell(pdfCell52);

                    //PdfPCell pdfCell53 = new PdfPCell(new Phrase(new Chunk("Formalización de la Alianza Estratégica”.", ARIAL8bLACK)));
                    ////pdfCell2.BackgroundColor = BaseColor.BLACK;
                    //pdfCell53.HorizontalAlignment = Element.ALIGN_LEFT;
                    //pdfCell53.VerticalAlignment = Element.ALIGN_LEFT;
                    //pdfCell31.Border = 0;
                    //pdfCuerpo55.AddCell(pdfCell53);

                    StringBuilder str = new StringBuilder();

                    str.Append("Al respecto , en calidad de 'Beneficiarios' por mutuo acuerdo deciden participar en el 3er concurso de los Servicios de Extension Agraria SEAR 2022");
                    str.Append("formalizan su participación a través de la propuesta "+extensionistadatos.vNombrePropuesta+" , para cuyo efecto proceden a suscribir la presente Acta.");
                    //str.Append("su participación a través de la propuesta “……………………(Nombre de la propuesta)");
                    //str.Append("…………………..”, para cuyo efecto proceden a suscribir la presente Acta.");

                    PdfPCell pdfCell54 = new PdfPCell(new Phrase(new Chunk(str.ToString(), ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell54.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell54.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell54.Border = 0;
                    pdfCuerpo55.AddCell(pdfCell54);

                    //PdfPCell pdfCell56 = new PdfPCell(new Phrase(new Chunk("Sin haber otro punto a tratar y leída esta acta, se levantó la sesión, siendo las ………………………. horas del mismo día, los presentes firmaron en señal de conformidad.", ARIAL8)));
                    ////pdfCell2.BackgroundColor = BaseColor.BLACK;
                    //pdfCell56.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    //pdfCell56.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    //pdfCell31.Border = 0;
                    //pdfCuerpo55.AddCell(pdfCell56);

                    document.Add(pdfCuerpo55);

                    document.Add(new Paragraph("\n"));
                    document.Add(new Paragraph("\n"));
                    document.Add(new Paragraph("\n"));

                    PdfPTable pdfCuerpo66 = new PdfPTable(1);

                    float[] widths6 = new float[] { 100f };

                    pdfCuerpo66.SetWidths(widths6);
                    pdfCuerpo66.WidthPercentage = 100f;
                    pdfCuerpo66.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                    PdfPCell pdfCell61 = new PdfPCell(new Phrase(new Chunk("______________________________________", ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell61.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell61.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell61.Border = 0;
                    pdfCuerpo66.AddCell(pdfCell61);

                    PdfPCell pdfCell62 = new PdfPCell(new Phrase(new Chunk("Nombre y Apellido:"+ extensionistadatos.vNombres + " " + extensionistadatos.vApemat + " " + extensionistadatos.vApepat, ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell62.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell62.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell62.Border = 0;
                    pdfCuerpo66.AddCell(pdfCell62);

                    PdfPCell pdfCell63 = new PdfPCell(new Phrase(new Chunk("Proveedor del SEAR – Extensionista", ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell63.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell63.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell63.Border = 0;
                    pdfCuerpo66.AddCell(pdfCell63);

                    document.Add(pdfCuerpo66);

                    document.Close();

                    byte[] buffer = memoryStream.ToArray();
                    var contentLength = buffer.Length;
                    var result = Request.CreateResponse(HttpStatusCode.OK);
                    result.Content = new StreamContent(new MemoryStream(buffer));
                    result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = Guid.NewGuid().ToString() + "_" + DateTime.Now.ToShortDateString() + ".pdf"
                    };
                    return result;
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
        public HttpResponseMessage GenerarPdfAlianzaEstrategica(Extensionista extensionista)
        {

            using (MemoryStream memoryStream = new MemoryStream())
            {
                try
                {
                    Document document = new Document(PageSize.A4.Rotate(), 40, 40, 100, 50);
                    PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);
                                        
                    //writer.PageEvent = new ITextEventsRetencion(HttpContext.Current.Server.MapPath("~/Image") + "/agrorural.jpg", str.ToString());                                        
                    document.Open();

                    string path = HttpContext.Current.Server.MapPath("~/Content/Images");
                    string imageURL = path + "/logoagrorural2.png";
                    iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imageURL);
                    //Resize image depend upon your need
                    jpg.ScaleToFit(150, 50);
                    //Give space before image
                    //jpg.SpacingBefore = 10f;
                    //Give some space after the image
                    //                    jpg.SpacingAfter = 10f;
                    jpg.Alignment = Element.ALIGN_RIGHT;
                    //jpg.SetAbsolutePosition(420, 770);
                    jpg.SetAbsolutePosition(650, 525);

                    document.Add(jpg);

                    string imageURL1 = path + "/logomdar.jpg";
                    iTextSharp.text.Image jpg1 = iTextSharp.text.Image.GetInstance(imageURL1);
                    //Resize image depend upon your need
                    jpg1.ScaleToFit(150, 50);
                    //Give space before image
                    //jpg.SpacingBefore = 10f;
                    //Give some space after the image
                    //                    jpg.SpacingAfter = 10f;
                    jpg1.Alignment = Element.ALIGN_RIGHT;
                    jpg1.SetAbsolutePosition(40,525);
                    document.Add(jpg1);

                    PdfPTable tableTitutlo = new PdfPTable(1);
                    float[] widths = new float[] { 100f };
                    tableTitutlo.SetWidths(widths);
                    PdfPCell pdfCell2 = new PdfPCell(new Phrase(new Chunk("Anexo N° 02. Acta de Alianza estratégica", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell2.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell2.VerticalAlignment = Element.ALIGN_CENTER;
                    pdfCell2.Border = 0;
                    tableTitutlo.AddCell(pdfCell2);
                    document.Add(tableTitutlo);

                    document.Add(new Paragraph("\n"));

                    ExtensionistaDL extensionistaDL = new ExtensionistaDL();

                    Extensionista extensionistadatos = extensionistaDL.ListarExtensionistaPorCodigo(extensionista);

                    PdfPTable pdfCuerpo1 = new PdfPTable(1);
                    float[] widths1 = new float[] { 150f };
                    pdfCuerpo1.SetWidths(widths1);
                    pdfCuerpo1.WidthPercentage = 100f;
                    pdfCuerpo1.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


                    PdfPCell pdfCell21 = new PdfPCell(new Phrase(new Chunk("En la ciudad de "+ extensionistadatos .vNomDistrito+ ", Provincia de "+ extensionistadatos.vNomProvincia+ ", y Departamento de "+ extensionistadatos.vNomDepartamento+ "; siendo las ………….. horas del día ………… del mes de ……………. del año 2022, los señores productores:", ARIAL8)));

                    pdfCell21.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell21.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell21.Border = 0;

                    pdfCuerpo1.AddCell(pdfCell21);
                                        
                    document.Add(pdfCuerpo1);

                    document.Add(new Paragraph("\n"));

                    ActaAlianzaEstrategicaDL actaAlianzaEstrategicaDL = new ActaAlianzaEstrategicaDL();
                    
                    // productores

                    PdfPTable pdfCuerpo33 = new PdfPTable(9);
                    float[] widths123 = new float[] { 10f,70f, 20F, 25f, 15f, 20f, 55f,55f, 25f };

                    pdfCuerpo33.SetWidths(widths123);
                    pdfCuerpo33.WidthPercentage = 100f;
                    //pdfCuerpo33.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                    PdfPCell pdfCell31 = new PdfPCell(new Phrase(new Chunk("N°", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell31.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell31.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell31.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell31);

                    PdfPCell pdfCell32 = new PdfPCell(new Phrase(new Chunk("Apellidos y Nombres", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell32.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell32.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell32.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell32);

                    PdfPCell pdfCell33 = new PdfPCell(new Phrase(new Chunk("DNI", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell33.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell33.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell33.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell33);

                    PdfPCell pdfCell34 = new PdfPCell(new Phrase(new Chunk("Celular", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell34.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell34.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell34.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell34);

                    PdfPCell pdfCell35 = new PdfPCell(new Phrase(new Chunk("Edad", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell35.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell35.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell35.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell35);

                    PdfPCell pdfCell36 = new PdfPCell(new Phrase(new Chunk("Sexo", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell36.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell36.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell36.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell36);

                    PdfPCell pdfCell37 = new PdfPCell(new Phrase(new Chunk("Nombre de la organización", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell37.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell37.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell37.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell37);


                    PdfPCell pdfCell391 = new PdfPCell(new Phrase(new Chunk("Recibio capacitacion y/o asistencia técnica", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell391.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell391.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell38.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell391);

                    PdfPCell pdfCell38 = new PdfPCell(new Phrase(new Chunk("Firma", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell38.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell38.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell38.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell38);                   
                                        
                    List<Productor> lista = new List<Productor>();
                    Productor productor = new Productor();

                    productor.iCodExtensionista = extensionista.iCodExtensionista;
                    productor.piPageSize = 40;
                    productor.piCurrentPage = 1;
                    productor.pvSortColumn = "iCodProductor";
                    productor.pvSortOrder = "asc";
                    productor.iPerteneceOrganizacion = 1;

                    lista = actaAlianzaEstrategicaDL.ListarProductor(productor);
                    int i = 0;
                    foreach (Productor item in lista)
                    {
                        PdfPCell pdfCell39 = new PdfPCell(new Phrase(new Chunk((i + 1).ToString(), ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell39.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell39.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell39);

                        PdfPCell pdfCell40 = new PdfPCell(new Phrase(new Chunk(item.vApellidosNombres, ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell40.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell40.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell40);

                        PdfPCell pdfCell41 = new PdfPCell(new Phrase(new Chunk(item.vDni, ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell41.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell41.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell41);

                        PdfPCell pdfCell45 = new PdfPCell(new Phrase(new Chunk(item.vCelular, ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell45.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell45.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell45);


                        PdfPCell pdfCell42 = new PdfPCell(new Phrase(new Chunk(item.iEdad.ToString(), ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell42.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell42.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell42);

                        PdfPCell pdfCell43 = new PdfPCell(new Phrase(new Chunk(item.iSexo==1 ? "Masculino":"Femenino", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell43.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell43.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell43);

                        PdfPCell pdfCell44 = new PdfPCell(new Phrase(new Chunk(item.vNombreOrganizacion, ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell44.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell44.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell44);

                        PdfPCell pdfCell46 = new PdfPCell(new Phrase(new Chunk(item.iRecibioCapacitacion==1 ? "SI":"NO", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell46.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell46.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell46);

                        PdfPCell pdfCell47 = new PdfPCell(new Phrase(new Chunk("", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell47.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell47.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell47);
                        i++;
                    }
                                        
                    document.Add(pdfCuerpo33);

                    //document.Add(new Paragraph("\n"));

                    PdfPTable pdfCuerpo55 = new PdfPTable(1);

                    float[] widths5 = new float[] { 100f};

                    pdfCuerpo55.SetWidths(widths5);
                    pdfCuerpo55.WidthPercentage = 100f;

                    PdfPCell pdfCell51 = new PdfPCell(new Phrase(new Chunk("En calidad de “Beneficiarios” por mutuo acuerdo deciden participar en el Servicios de Extensión Agraria Rural SEAR 2022.", ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell51.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell51.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell51.Border = 0;
                    pdfCuerpo55.AddCell(pdfCell51);

                 

                    PdfPCell pdfCell52 = new PdfPCell(new Phrase(new Chunk("Así mismo, el Sr. "+ extensionistadatos.vNombres+" "+extensionistadatos.vApemat+" "+extensionistadatos.vApepat + " Identificado con número de DNI Nº "+extensionistadatos.vDni+" y RUC Nº "+ extensionistadatos.vRuc+".decide participar en calidad de “Proveedor del SEAR – Extensionista”.", ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell52.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell52.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell52.Border = 0;
                    pdfCuerpo55.AddCell(pdfCell52);
                                        
                    PdfPCell pdfCell53 = new PdfPCell(new Phrase(new Chunk("Formalización de la Alianza Estratégica”.", ARIAL8bLACK)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell53.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell53.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell53.Border = 0;
                    pdfCuerpo55.AddCell(pdfCell53);

                    StringBuilder str = new StringBuilder();

                    str.Append("Al respecto, “los beneficiarios” y el “Proveedor del SEAR”, luego de evaluar la importancia de ");
                    str.Append("participar en el 3.er concurso de los Servicios de Extensión Agraria Rural SEAR 2022, formalizan");
                    str.Append("su participación a través de la propuesta " + extensionistadatos.vNombrePropuesta);
                    str.Append(", para cuyo efecto proceden a suscribir la presente Acta.");

                    PdfPCell pdfCell54 = new PdfPCell(new Phrase(new Chunk(str.ToString(), ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell54.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell54.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell54.Border = 0;
                    pdfCuerpo55.AddCell(pdfCell54);

                    PdfPCell pdfCell56 = new PdfPCell(new Phrase(new Chunk("Sin haber otro punto a tratar y leída esta acta, se levantó la sesión, siendo las ………………………. horas del mismo día, los presentes firmaron en señal de conformidad.", ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell56.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell56.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell56.Border = 0;
                    pdfCuerpo55.AddCell(pdfCell56);

                    PdfPCell pdfCell58 = new PdfPCell(new Phrase(new Chunk("-", ARIAL8white)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell58.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell58.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell58.Border = 0;
                    pdfCuerpo55.AddCell(pdfCell58);

                    PdfPCell pdfCell59 = new PdfPCell(new Phrase(new Chunk("-", ARIAL8white)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell59.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell59.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell59.Border = 0;
                    pdfCuerpo55.AddCell(pdfCell59);

                    // Representantes inicio
                    ActaAlianzaEstrategicaDL alianzaEstrategicaDL = new ActaAlianzaEstrategicaDL();
                    Productor productorparametro = new Productor();
                    productorparametro.iCodExtensionista = extensionista.iCodExtensionista;
                    List<Productor> listadrepresentes = alianzaEstrategicaDL.ListarRepresentantes(productorparametro);


                    PdfPTable pdfCuerpoparticipantes = new PdfPTable(3);

                    float[] widths7 = new float[] { 33f,33f,33f };

                    pdfCuerpoparticipantes.SetWidths(widths7);
                    pdfCuerpoparticipantes.WidthPercentage = 100f;

                    int cantidad_representantes = listadrepresentes.Count;


                    foreach (Productor item in listadrepresentes)
                    {
                        PdfPCell pdfCellPartItem = new PdfPCell(new Phrase(new Chunk("____________________________________________"+Chunk.NEWLINE+"Nombre y Apellido: " +item.vApellidosNombres +Chunk.NEWLINE + "Representante de la Organizacion "+item.vNombreOrganizacion, ARIAL8) ));

                        pdfCellPartItem.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        pdfCellPartItem.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                        pdfCellPartItem.Border = 0;
                        pdfCuerpoparticipantes.AddCell(pdfCellPartItem);
                    }

                    if(cantidad_representantes==1)
                    {
                        PdfPCell pdfCellPartItemparche1 = new PdfPCell(new Phrase(new Chunk("")));

                        pdfCellPartItemparche1.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        pdfCellPartItemparche1.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                        pdfCellPartItemparche1.Border = 0;
                        pdfCuerpoparticipantes.AddCell(pdfCellPartItemparche1);

                        PdfPCell pdfCellPartItemparche2 = new PdfPCell(new Phrase(new Chunk("")));

                        pdfCellPartItemparche2.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        pdfCellPartItemparche2.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                        pdfCellPartItemparche2.Border = 0;
                        pdfCuerpoparticipantes.AddCell(pdfCellPartItemparche2);
                    }
                    if(cantidad_representantes==2)
                    {
                        PdfPCell pdfCellPartItemparche3 = new PdfPCell(new Phrase(new Chunk("")));

                        pdfCellPartItemparche3.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                        pdfCellPartItemparche3.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                        pdfCellPartItemparche3.Border = 0;
                        pdfCuerpoparticipantes.AddCell(pdfCellPartItemparche3);
                    }

                    
                    PdfPCell pdfCell57 = new PdfPCell(new Phrase(new Chunk("Participantes")));

                    pdfCell57.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell57.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell57.Border = 0;
                    pdfCell57.AddElement(pdfCuerpoparticipantes);
                    pdfCuerpo55.AddCell(pdfCell57);

                    //document.Add(new Paragraph("\n"));
                    //document.Add(new Paragraph("\n"));
                    //document.Add(new Paragraph("\n"));

                    // Representantes fin

                    document.Add(pdfCuerpo55);

                    //document.Add(new Paragraph("\n"));
                    //document.Add(new Paragraph("\n"));
                    document.Add(new Paragraph("\n"));

                    PdfPTable pdfCuerpo66 = new PdfPTable(1);

                    float[] widths6 = new float[] { 100f };

                    pdfCuerpo66.SetWidths(widths6);
                    pdfCuerpo66.WidthPercentage = 100f;
                    pdfCuerpo66.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                    PdfPCell pdfCell61 = new PdfPCell(new Phrase(new Chunk("________________________________________", ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell61.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell61.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell61.Border = 0;
                    pdfCuerpo66.AddCell(pdfCell61);

                    PdfPCell pdfCell62 = new PdfPCell(new Phrase(new Chunk("Nombre y Apellido: " + extensionistadatos.vNombres + " " + extensionistadatos.vApemat + " " + extensionistadatos.vApepat , ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell62.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell62.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell62.Border = 0;
                    pdfCuerpo66.AddCell(pdfCell62);

                    PdfPCell pdfCell63 = new PdfPCell(new Phrase(new Chunk("Proveedor del SEAR – Extensionista", ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell63.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfCell63.VerticalAlignment = Element.ALIGN_LEFT;
                    pdfCell63.Border = 0;
                    pdfCuerpo66.AddCell(pdfCell63);

                    document.Add(pdfCuerpo66);

                    document.Close();

                    byte[] buffer = memoryStream.ToArray();
                    var contentLength = buffer.Length;
                    var result = Request.CreateResponse(HttpStatusCode.OK);
                    result.Content = new StreamContent(new MemoryStream(buffer));
                    result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = Guid.NewGuid().ToString() + "_" + DateTime.Now.ToShortDateString() + ".pdf"
                    };
                    return result;
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }

        public HttpResponseMessage GenerarPdfActa(Extensionista extensionista)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                try
                {
                    Document document = new Document(PageSize.A4, 40, 40, 100, 50);
                    PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);

                    //writer.PageEvent = new ITextEventsRetencion(HttpContext.Current.Server.MapPath("~/Image") + "/agrorural.jpg", str.ToString());                                        
                    document.Open();

                    string path = HttpContext.Current.Server.MapPath("~/Content/Images");
                    string imageURL = path + "/logoagrorural2.png";
                    iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imageURL);
                    //Resize image depend upon your need
                    jpg.ScaleToFit(140, 40);
                    //Give space before image
                    //jpg.SpacingBefore = 10f;
                    //Give some space after the image
                    //                    jpg.SpacingAfter = 10f;
                    jpg.Alignment = Element.ALIGN_RIGHT;
                    jpg.SetAbsolutePosition(420, 770);

                    document.Add(jpg);

                    string imageURL1 = path + "/logomdar.jpg";
                    iTextSharp.text.Image jpg1 = iTextSharp.text.Image.GetInstance(imageURL1);
                    //Resize image depend upon your need
                    jpg1.ScaleToFit(150, 50);
                    //Give space before image
                    //jpg.SpacingBefore = 10f;
                    //Give some space after the image
                    //                    jpg.SpacingAfter = 10f;
                    jpg1.Alignment = Element.ALIGN_RIGHT;
                    jpg1.SetAbsolutePosition(20, 770);
                    document.Add(jpg1);

                    PdfPTable tableTituto = new PdfPTable(1);
                    float[] widths = new float[] { 100f };
                    tableTituto.SetWidths(widths);
                    PdfPCell pdfCell2 = new PdfPCell(new Phrase(new Chunk("Anexo N° 03. Acta de compromiso de la organizacion de productores", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell2.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell2.VerticalAlignment = Element.ALIGN_CENTER;
                    pdfCell2.Border = 0;
                    tableTituto.AddCell(pdfCell2);
                    document.Add(tableTituto);

                    //document.Add(new Paragraph("\n"));

                    PdfPTable tableTitulo2 = new PdfPTable(1);
                    float[] widths2 = new float[] { 100f };
                    tableTitulo2.SetWidths(widths2);
                    PdfPCell pdfCell21 = new PdfPCell(new Phrase(new Chunk("SERVICIO:[Nomnbre de la Ficha Tecnica de Servicios de Extension]", ARIAL13)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell21.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell21.VerticalAlignment = Element.ALIGN_CENTER;
                    pdfCell21.Border = 0;
                    tableTitulo2.AddCell(pdfCell21);

                    StringBuilder strtexto = new StringBuilder();

                    strtexto.Append("En la comunidad/sector/localidad de ......................................Distrito de..........................");
                    strtexto.Append("Provincia de................. y Departamento de....................................");
                    strtexto.Append("siendo las horas del dia ...... del mes ..... del año 2022 en las instalaciones de");
                    strtexto.Append("......................................se reunieron los siguientes productores:");

                    PdfPCell pdfCell22 = new PdfPCell(new Phrase(new Chunk(strtexto.ToString(), ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell22.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell22.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell22.Border = 0;
                    tableTitulo2.AddCell(pdfCell22);
                    document.Add(tableTitulo2);

                    document.Add(new Paragraph("\n"));

                    // productores

                    PdfPTable pdfCuerpo33 = new PdfPTable(7);
                    float[] widths123 = new float[] { 15f, 80f, 25F, 25f, 30f, 30f, 40 };

                    pdfCuerpo33.SetWidths(widths123);
                    pdfCuerpo33.WidthPercentage = 100f;
                    //pdfCuerpo33.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                    PdfPCell pdfCell31 = new PdfPCell(new Phrase(new Chunk("N°", ARIAL13)));
                    pdfCell31.BackgroundColor = BaseColor.LIGHT_GRAY;
                    pdfCell31.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell31.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell31.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell31);

                    PdfPCell pdfCell32 = new PdfPCell(new Phrase(new Chunk("Apellidos y Nombres", ARIAL13)));
                    pdfCell32.BackgroundColor = BaseColor.LIGHT_GRAY;
                    pdfCell32.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell32.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell32.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell32);

                    PdfPCell pdfCell33 = new PdfPCell(new Phrase(new Chunk("DNI", ARIAL13)));
                    pdfCell33.BackgroundColor = BaseColor.LIGHT_GRAY;
                    pdfCell33.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell33.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell33.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell33);

                    PdfPCell pdfCell34 = new PdfPCell(new Phrase(new Chunk("Celular", ARIAL13)));
                    pdfCell34.BackgroundColor = BaseColor.LIGHT_GRAY;
                    pdfCell34.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell34.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell34.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell34);


                    PdfPCell pdfCell35 = new PdfPCell(new Phrase(new Chunk("aaaaaaa", ARIAL13)));
                    pdfCell35.BackgroundColor = BaseColor.LIGHT_GRAY;
                    pdfCell35.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell35.VerticalAlignment = Element.ALIGN_CENTER;
                    pdfCell35.Colspan = 2;
                    pdfCell35.Padding = 0;
                    // tabla incio cabecera Unidad productiva
                    
                    PdfPTable pdfCuerpo51 = new PdfPTable(2);
                    float[] widths51 = new float[] { 10f, 10f };

                    pdfCuerpo51.SetWidths(widths51);
                    pdfCuerpo51.WidthPercentage = 100f;

                    PdfPCell pdfCell511 = new PdfPCell(new Phrase(new Chunk("UNIDAD PRODUCTIVA", ARIAL13)));
                    //pdfCell511.BackgroundColor = BaseColor.tra;
                    pdfCell511.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell511.VerticalAlignment = Element.ALIGN_CENTER;
                    pdfCell511.Colspan = 2;
                    pdfCuerpo51.AddCell(pdfCell511);

                    //PdfPCell pdfCell512 = new PdfPCell(new Phrase(new Chunk("2", ARIAL13)));
                    ////pdfCell512.BackgroundColor = BaseColor.LIGHT_GRAY;
                    //pdfCell512.HorizontalAlignment = Element.ALIGN_CENTER;
                    //pdfCell512.VerticalAlignment = Element.ALIGN_CENTER;

                    //pdfCuerpo51.AddCell(pdfCell512);

                    PdfPCell pdfCell513 = new PdfPCell(new Phrase(new Chunk("Hectáreas (has)", ARIAL13)));
                    //pdfCell511.BackgroundColor = BaseColor.tra;
                    pdfCell513.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell513.VerticalAlignment = Element.ALIGN_CENTER;

                    pdfCuerpo51.AddCell(pdfCell513);

                    PdfPCell pdfCell514 = new PdfPCell(new Phrase(new Chunk("Cabezas (und)", ARIAL13)));
                    //pdfCell512.BackgroundColor = BaseColor.LIGHT_GRAY;
                    pdfCell514.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell514.VerticalAlignment = Element.ALIGN_CENTER;

                    pdfCuerpo51.AddCell(pdfCell514);


                    // tabla fin cabecera Unidad productiva

                    pdfCell35.AddElement(pdfCuerpo51);

                    //pdfCell35.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell35);

                    //PdfPCell pdfCell36 = new PdfPCell(new Phrase(new Chunk("", ARIAL13)));
                    ////pdfCell2.BackgroundColor = BaseColor.BLACK;
                    //pdfCell36.HorizontalAlignment = Element.ALIGN_CENTER;
                    //pdfCell36.VerticalAlignment = Element.ALIGN_CENTER;
                    ////pdfCell36.Border = 0;
                    //pdfCuerpo33.AddCell(pdfCell36);



                    PdfPCell pdfCell37 = new PdfPCell(new Phrase(new Chunk("Firma", ARIAL13)));
                    pdfCell37.BackgroundColor = BaseColor.LIGHT_GRAY;
                    pdfCell37.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCell37.VerticalAlignment = Element.ALIGN_CENTER;
                    //pdfCell37.Border = 0;
                    pdfCuerpo33.AddCell(pdfCell37);

                    for (int i = 0; i <= 39; i++)
                    {
                        PdfPCell pdfCell39 = new PdfPCell(new Phrase(new Chunk((i + 1).ToString(), ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell39.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell39.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell39);

                        PdfPCell pdfCell40 = new PdfPCell(new Phrase(new Chunk("", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell40.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell40.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell40);

                        PdfPCell pdfCell41 = new PdfPCell(new Phrase(new Chunk("", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell41.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell41.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell41);

                        PdfPCell pdfCell45 = new PdfPCell(new Phrase(new Chunk("", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell45.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell45.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell45);

                        PdfPCell pdfCell42 = new PdfPCell(new Phrase(new Chunk("", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell42.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell42.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell42);

                        PdfPCell pdfCell43 = new PdfPCell(new Phrase(new Chunk("", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell43.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell43.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell43);

                        PdfPCell pdfCell44 = new PdfPCell(new Phrase(new Chunk("", ARIAL8)));
                        //pdfCell2.BackgroundColor = BaseColor.BLACK;
                        pdfCell44.HorizontalAlignment = Element.ALIGN_CENTER;
                        pdfCell44.VerticalAlignment = Element.ALIGN_CENTER;
                        //pdfCell38.Border = 0;
                        pdfCuerpo33.AddCell(pdfCell44);                                                
                    }

                    document.Add(pdfCuerpo33);

                    PdfPTable table3= new PdfPTable(1);
                    float[] widths3 = new float[] { 100f };
                    table3.SetWidths(widths3);

                    StringBuilder strtexto1 = new StringBuilder();

                    strtexto1.Append("al respecto los mencionados productores por mutuo acuerdo deciden participar");
                    strtexto1.Append(" como [Nomnbres de la Organizacion de productores] en el 3° Concurso de Servicios de Extension ");
                    strtexto1.Append(" Agraria Rural SEAR - 2022, a fin de acceder a los servicios de capacitación y asistencia tecnica del PI ");
                    strtexto1.Append("'MEJORAMIENTO DE LAS CAPACIDADES DE LAS DIRECCIONES REGIONALES AGRARIAS Y AGENCIAS AGRARIAS EN 11 DEPARTAMENTOS'");
                    strtexto1.Append(" con CUI 2516447, En señal de conformidad firman el Extensionista responsable de presentar la propuesta ");
                    strtexto1.Append(" [Nombre de la ficha Tecnica de Servicios de Extension] y a la junta directiva de la organizacion de productores");

                    PdfPCell pdfCell131 = new PdfPCell(new Phrase(new Chunk(strtexto1.ToString(), ARIAL8)));
                    //pdfCell2.BackgroundColor = BaseColor.BLACK;
                    pdfCell131.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell131.VerticalAlignment = Element.ALIGN_JUSTIFIED;
                    pdfCell131.Border = 0;
                    table3.AddCell(pdfCell131);

                    document.Add(table3);

                    document.Close();

                    byte[] buffer = memoryStream.ToArray();
                    var contentLength = buffer.Length;
                    var result = Request.CreateResponse(HttpStatusCode.OK);
                    result.Content = new StreamContent(new MemoryStream(buffer));
                    result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = Guid.NewGuid().ToString() + "_" + DateTime.Now.ToShortDateString() + ".pdf"
                    };
                    return result;

                }
                catch(Exception ex)
                {
                    throw ex;
                }
             }
        }
        private Boolean BuscarSeleccionadoRequisito(List<ListaChequeoRequisitos> lista,int idrequisito)
        {
            Boolean respuesta = false;
            foreach (ListaChequeoRequisitos item in lista)
            {
                if(item.iCodRequisito==idrequisito)
                {
                    respuesta=item.bCumple;
                    break;
                }
            }
            return respuesta;
        }
        private PdfPTable GenerateTable(string uno, string dos)
        {
            PdfPTable tablegenerate = new PdfPTable(2);
            float[] withGenerate = new float[] { 150f, 150f };
            tablegenerate.SetWidths(withGenerate);
            tablegenerate.WidthPercentage = 60f;
            PdfPCell cellGenerate = new PdfPCell(new Phrase(new Chunk(uno, ARIAL8bLACK)));
            cellGenerate.HorizontalAlignment = Element.ALIGN_CENTER;
            cellGenerate.VerticalAlignment = Element.ALIGN_CENTER;
            cellGenerate.BackgroundColor = BaseColor.LIGHT_GRAY;
            tablegenerate.AddCell(cellGenerate);
            PdfPCell cellGenerate1 = new PdfPCell(new Phrase(new Chunk(dos, ARIAL8bLACK)));
            cellGenerate1.HorizontalAlignment = Element.ALIGN_CENTER;
            cellGenerate1.VerticalAlignment = Element.ALIGN_CENTER;
            tablegenerate.AddCell(cellGenerate1);
            return tablegenerate;
        }
        private PdfPTable GenerateTableUnaFila(string uno,Font font,int titulo,int alineacion,int imagen,Boolean estado)
        {
            PdfPTable tablegenerate = new PdfPTable(1);
            PdfPCell cellGenerate = new PdfPCell();
            //tablegenerate.WidthPercentage = 100f;
            string path = HttpContext.Current.Server.MapPath("~/Content/Images");

            if (imagen==1)
            {
                string imageURL = "";
                if(estado==false)
                {
                    imageURL = path+"/radionoseleccionado.jpg";
                }
                else
                {
                    imageURL = path + "/radioseleccionado.jpg";
                }                
                
                iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imageURL);
                //Resize image depend upon your need
                jpg.ScaleToFit(15, 15);
                //Give space before image
                jpg.SpacingBefore = 10f;
                //Give some space after the image
                jpg.SpacingAfter = 10f;
                jpg.Alignment = Element.ALIGN_CENTER;
                cellGenerate.AddElement(jpg);
            }
            else
            {
                cellGenerate.AddElement(new Phrase(new Chunk(uno, font)));
            }
            
            if(alineacion==1)
            {
                cellGenerate.HorizontalAlignment = Element.ALIGN_CENTER;
                cellGenerate.VerticalAlignment = Element.ALIGN_CENTER;
            }
            if (alineacion == 2)
            {
                cellGenerate.HorizontalAlignment = Element.ALIGN_JUSTIFIED_ALL;
                cellGenerate.VerticalAlignment = Element.ALIGN_JUSTIFIED_ALL;
            }

            if (titulo==1)
            {
                cellGenerate.BackgroundColor = BaseColor.LIGHT_GRAY;
            }                        
            tablegenerate.AddCell(cellGenerate);
            return tablegenerate;
        }

        private PdfPTable GenerateTableDosFila(string uno, string dos, string cuatro, string cinco)
        {
            PdfPTable tablegenerate = new PdfPTable(2);
            float[] withGenerate = new float[] { 150f, 150f };
            tablegenerate.SetWidths(withGenerate);
            tablegenerate.WidthPercentage = 70f;
            PdfPCell cellGenerate = new PdfPCell(new Phrase(new Chunk(uno, ARIAL8bLACK)));
            cellGenerate.HorizontalAlignment = Element.ALIGN_CENTER;
            cellGenerate.VerticalAlignment = Element.ALIGN_CENTER;
            cellGenerate.BackgroundColor = BaseColor.LIGHT_GRAY;
            tablegenerate.AddCell(cellGenerate);

            PdfPCell cellGenerate1 = new PdfPCell(new Phrase(new Chunk(dos, ARIAL8bLACK)));
            cellGenerate1.HorizontalAlignment = Element.ALIGN_CENTER;
            cellGenerate1.VerticalAlignment = Element.ALIGN_CENTER;
            tablegenerate.AddCell(cellGenerate1);

            PdfPCell cellGenerate123 = new PdfPCell(new Phrase(new Chunk(cuatro, ARIAL8bLACK)));
            cellGenerate123.HorizontalAlignment = Element.ALIGN_CENTER;
            cellGenerate123.VerticalAlignment = Element.ALIGN_CENTER;
            cellGenerate123.BackgroundColor = BaseColor.LIGHT_GRAY;
            tablegenerate.AddCell(cellGenerate123);

            PdfPCell cellGenerate1234 = new PdfPCell(new Phrase(new Chunk(cinco, ARIAL8bLACK)));
            cellGenerate1234.HorizontalAlignment = Element.ALIGN_CENTER;
            cellGenerate1234.VerticalAlignment = Element.ALIGN_CENTER;
            tablegenerate.AddCell(cellGenerate1234);

            return tablegenerate;
        }

        //public class Productor
        //{
        //    public int numero { get; set; }
        //    public string NombresyApellidos { get; set; }
        //    public string dni { get; set; }
        //    public int edad { get; set; }
        //    public int sexo { get; set; }

        //    public string celular { get; set; }

        //    public string nombreorg { get; set; }

        //}
        public class ITextEventsRetencion : PdfPageEventHelper
        {
            PdfContentByte cb;
            PdfTemplate headerTemplate;
            BaseFont bf = null;
            DateTime PrintTime = DateTime.Now;
            #region Fields
            private string _imageUrl;
            private string _TextoMain;

            public ITextEventsRetencion(string imageUrl, string TextoMain)
            {
                _imageUrl = imageUrl;
                _TextoMain = TextoMain;
            }
            #endregion
            iTextSharp.text.Font baseFontNormal = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            iTextSharp.text.Font baseFontNormalNormal = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);

            iTextSharp.text.Font ARIAL10 = new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 9, iTextSharp.text.Font.BOLD, BaseColor.BLACK);

            public override void OnOpenDocument(PdfWriter writer, Document document)
            {
                try
                {
                    PrintTime = DateTime.Now;
                    bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb = writer.DirectContent;
                    headerTemplate = cb.CreateTemplate(100, 100);
                }
                catch (DocumentException)
                {

                }
                catch (IOException)
                {

                }
            }

            public override void OnEndPage(PdfWriter writer, Document document)
            {
                base.OnEndPage(writer, document);

                float[] anchoDeColumnas = new float[] { 280f, 80f };
                PdfPTable pdfTab = new PdfPTable(2);
                pdfTab.SetWidths(anchoDeColumnas);

                iTextSharp.text.Image imagenBanner = iTextSharp.text.Image.GetInstance(_imageUrl);
                imagenBanner.ScaleToFit(anchoDeColumnas[0], 50f);
                imagenBanner.Alignment = Element.ALIGN_MIDDLE;

                PdfPCell pdfCell1 = new PdfPCell(imagenBanner);
                pdfCell1.Padding = 5f;
                pdfCell1.HorizontalAlignment = Element.ALIGN_CENTER;
                pdfCell1.VerticalAlignment = Element.ALIGN_CENTER;

                PdfPCell pdfCell2 = new PdfPCell(new Phrase(_TextoMain, ARIAL10));
                pdfCell2.Padding = 5f;
                pdfCell2.BorderColor = new BaseColor(System.Drawing.Color.Black);
                pdfCell2.HorizontalAlignment = Element.ALIGN_LEFT;
                pdfCell2.VerticalAlignment = Element.ALIGN_CENTER;

                pdfTab.AddCell(pdfCell2);
                pdfTab.AddCell(pdfCell1);

                pdfTab.TotalWidth = document.PageSize.Width - 60f;
                pdfTab.WidthPercentage = 100;

                pdfTab.WriteSelectedRows(0, -1, 20, document.PageSize.Height - 30, writer.DirectContent);

                int pageN = writer.PageNumber;
                string text = $"Pagina {Convert.ToString(pageN)} de ";
                float len = bf.GetWidthPoint(text, 12);
                iTextSharp.text.Rectangle pageSize = document.PageSize;
                cb.SetRGBColorFill(100, 100, 100);
                cb.BeginText();
                cb.SetFontAndSize(bf, 12);
                cb.SetTextMatrix(document.RightMargin, pageSize.GetBottom(document.BottomMargin - 10));
                cb.ShowText(text);
                cb.EndText();
                cb.AddTemplate(headerTemplate, document.RightMargin + len, pageSize.GetBottom(document.BottomMargin - 10));
            }

            public override void OnCloseDocument(PdfWriter writer, Document document)
            {
                base.OnCloseDocument(writer, document);
                headerTemplate.BeginText();
                headerTemplate.SetFontAndSize(bf, 12);
                headerTemplate.SetTextMatrix(0, 0);
                headerTemplate.ShowText("" + (writer.PageNumber));
                headerTemplate.EndText();
            }
        }
    }

}
