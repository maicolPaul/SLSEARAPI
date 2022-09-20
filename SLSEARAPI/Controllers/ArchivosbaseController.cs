using SLSEARAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace SLSEARAPI.Controllers
{
    public class ArchivosbaseController : ApiController
    {
        protected string fileDafult;
        protected string server;
        protected string user;
        protected string password;
        protected string dominio;

        public ArchivosbaseController()
        {
            fileDafult = Convert.ToBase64String(File.ReadAllBytes(HttpContext.Current.Server.MapPath("~/Content/Images") + "\\no-file.png"));
            server = ConfigurationManager.AppSettings["FTPServer"].ToString();
            user = ConfigurationManager.AppSettings["FTPUser"].ToString();
            password = ConfigurationManager.AppSettings["FTPPassword"].ToString();
            dominio = ConfigurationManager.AppSettings["FTPDominio"].ToString();
        }

        public async Task<Archivo> ValidarSubirArchivosAsync(bool multiParse, string path = "", string extension = "")
        {
            if (!multiParse)
            {
                return new Archivo
                {
                    mensaje = "No es FormData",
                    validation = false
                };
            }
            if (string.IsNullOrEmpty(path))
            {
                return new Archivo
                {
                    mensaje = "Path del Archivo Vacio",
                    validation = false
                };
            }
            int count = 0;
            List<string> fileName = new List<string>();
            MultipartFormDataStreamProvider provider = new MultipartFormDataStreamProvider(HttpContext.Current.Server.MapPath("~/App_Data"));
            try
            {
                await Request.Content.ReadAsMultipartAsync(provider);
                if (provider.FileData.Count.Equals(0))
                {
                    return new Archivo
                    {
                        mensaje = "No se Mandaron Archivos",
                        validation = false
                    };
                }
                if (!string.IsNullOrEmpty(extension))
                {
                    List<string> list = ConvertstringToArray(extension, ',');
                    foreach (MultipartFileData fileDatum in provider.FileData)
                    {
                        string ext = Path.GetExtension(fileDatum.Headers.ContentDisposition.FileName.Replace("\"", "")).ToUpper();
                        if (list.Find((string item) => item.ToUpper().Equals(ext)) == null)
                        {
                            return new Archivo
                            {
                                mensaje = $"La Extension {ext} no esta Permitido",
                                validation = false
                            };
                        }
                    }
                }
                foreach (MultipartFileData fileDatum2 in provider.FileData)
                {
                    string text = $"{Guid.NewGuid().ToString()}___{fileDatum2.Headers.ContentDisposition.FileName.Replace("\"", "")}";
                    var x = fileDatum2.LocalFileName;
                    FtpFileUpload(x, text, path);
                    fileName.Add(path + text);
                    File.Delete(x);
                    count++;
                }
                return new Archivo
                {
                    fileNames = fileName,
                    mensaje = $"Se Subieron {count} de {provider.FileData.Count} Archivos",
                    validation = true
                };
            }
            catch (Exception ex)
            {
                return new Archivo()
                {
                    mensaje = $"{ex.Source} \n {ex.Message} \n {ex.StackTrace}\n",
                    validation = false
                };
            }
        }

        public Archivo FtpFileUpload(string fullPath, string filePath, string concatServerUrl)
        {
            FtpWebRequest request;
            try
            {
                string requestUriString = server + concatServerUrl + filePath;
                request = WebRequest.Create(requestUriString) as FtpWebRequest;
                request.Method = WebRequestMethods.Ftp.UploadFile;
                request.UseBinary = true;
                request.UsePassive = true;
                request.KeepAlive = true;
                request.Credentials = new NetworkCredential(user, password);
                request.ConnectionGroupName = "group";
                using (FileStream fs = File.OpenRead(fullPath))
                {
                    byte[] array = new byte[fs.Length];
                    fs.Read(array, 0, array.Length);
                    fs.Close();
                    Stream requestStream = request.GetRequestStream();
                    requestStream.Write(array, 0, array.Length);
                    requestStream.Flush();
                    requestStream.Close();
                }
                return new Archivo
                {
                    mensaje = "OK",
                    validation = true
                };
            }
            catch (Exception ex)
            {
                return new Archivo()
                {
                    mensaje = $"{ex.Source} \n {ex.Message} \n {ex.StackTrace}\n",
                    validation = false
                };
            }
        }

        private bool CheckIfFileExistsOnServer(string vFoto)
        {
            var request = (FtpWebRequest)WebRequest.Create(server + vFoto);
            request.Credentials = new NetworkCredential(user, password);
            request.Method = WebRequestMethods.Ftp.GetFileSize;
            FtpWebResponse ftpWebResponse;
            try
            {
                ftpWebResponse = (FtpWebResponse)request.GetResponse();
                return true;
            }
            catch (WebException ex)
            {
                ftpWebResponse = (FtpWebResponse)ex.Response;
                if (ftpWebResponse.StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable)
                {
                    return false;
                }
            }
            return false;
        }

        public List<string> ConvertstringToArray(string data, char split)
        {
            List<string> list = new List<string>();
            foreach (var z in data.Split(split))
                list.Add(z.ToUpper());
            return list;
        }

        public Archivo FTPaBase64(string vFile)
        {
            Archivo archivo = new Archivo();
            using (WebClient webClient = new WebClient())
            {
                webClient.Credentials = new NetworkCredential(user, password);
                if (!CheckIfFileExistsOnServer(vFile))
                {
                    archivo.file = Convert.FromBase64String(fileDafult);
                    archivo.encode64 = fileDafult;
                    archivo.validation = true;
                    archivo.mensaje = "No Existe El Archivo";
                    archivo.mineType = "image/png";
                    archivo.path = "No_Archivo.png";
                    return archivo;
                }
                byte[] C = webClient.DownloadData(server + vFile);
                archivo.file = C;
                archivo.encode64 = Convert.ToBase64String(C);
                archivo.validation = true;
                archivo.mensaje = "Se Obtuvo el Archivo";
                archivo.mineType = MineTypeMap.GetMimeType(Path.GetExtension(vFile));
                archivo.path = vFile.Split(new[] { "___" }, StringSplitOptions.None).Last();
                return archivo;
            }
        }

        private static string ValidateBase64EncodedString(string inputText)
        {
            string stringToValidate = inputText;
            stringToValidate = stringToValidate.Replace('-', '+'); // 62nd char of encoding
            stringToValidate = stringToValidate.Replace('_', '/'); // 63rd char of encoding
            switch (stringToValidate.Length % 4) // Pad with trailing '='s
            {
                case 0: break; // No pad chars in this case
                case 2: stringToValidate += "=="; break; // Two pad chars
                case 3: stringToValidate += "="; break; // One pad char
                default:
                    throw new System.Exception(
             "Illegal base64url string!");
            }

            return stringToValidate;
        }
    }
}

