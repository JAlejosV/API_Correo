using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using API_Correo.DBContext;
using API_Correo.DTO;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace API_Correo.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ListaERController : ControllerBase
    {
        private readonly InterfacesDBContext interfacesDB;
        private readonly IConfiguration config;
        public ListaERController(InterfacesDBContext _interfacesDB, IConfiguration _config)
        {
            interfacesDB = _interfacesDB;
            config = _config;
        }
        // GET: api/<ListaERController>
        //[HttpGet]
        //public IEnumerable<string> Get()
        //{
        //    return new string[] { "value1", "value2" };
        //}
        // GET: api/<ListaERController>
        [HttpGet]
        public IActionResult Get(bool EnvioCorreo)
        {
            var listaER = interfacesDB.Set<EROfisisDTO>().FromSqlRaw($"exec USP_LISTA_ER").ToList();
            var envio = false;
            if (EnvioCorreo && listaER.Count > 0)
            {
                //Crear archivo excel
                var stream = new MemoryStream();
                using (var package = new ExcelPackage(stream))
                {
                    var workSheet = package.Workbook.Worksheets.Add("ListaER");
                    workSheet.Cells.LoadFromCollection(listaER, true);
                    package.Save();
                }
                stream.Position = 0;
                var archivoExcel = stream.ToArray();
                string excelName = string.Format("ListaER-{0}.xlsx", DateTime.Now.ToString("yyyyMMddHHmmssfff"));

                //Enviar Correo
                Correo correo = new Correo();

                correo.CorreoEmisor.Mail = config.GetValue<string>("EnvioLogCorreo:CorreoEnvio:Mail");
                correo.CorreoEmisor.NameMail = config.GetValue<string>("EnvioLogCorreo:CorreoEnvio:NameMail");
                correo.CorreoEmisor.Password = config.GetValue<string>("EnvioLogCorreo:CorreoEnvio:Password");

                var CorreoDefectoPara = config.GetSection("EnvioLogCorreo:CorreoDefectoPara").Get<List<string>>();

                if (CorreoDefectoPara.Count > 0)
                {
                    foreach (var correoPara in CorreoDefectoPara)
                    {
                        correo.CorreosPara.Add(new StructMail
                        {
                            Mail = correoPara,
                            NameMail = string.Empty
                        });
                    }
                }

                var ConCopia = config.GetSection("EnvioLogCorreo:ConCopia").Get<List<string>>();
                if (ConCopia.Count > 0)
                {
                    foreach (var CorreoCC in ConCopia)
                    {
                        try
                        {
                            correo.CorreosCC.Add(new StructMail
                            {
                                Mail = CorreoCC,
                                NameMail = string.Empty
                            }
                            );
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }
                    }
                }

                correo.Adjuntos.Add(new Adjunto
                {
                    archivo = archivoExcel,
                    nombreArchivo = excelName
                });

                try
                {
                    correo.Asunto = $"Listado de ER";

                    string cuerpo = "Existen ER pendientes de rendir, en el excel adjunto se detallan las ER.";

                    Helpers.Helpers.ConstruirCorreoError(correo, cuerpo);


                    envio = Helpers.Helpers.EnviarCorreoElectronico(correo, true);
                }
                catch (Exception ex)
                {

                }
            }

            return Ok(new
            {
                listaER,
                correo = envio
            });
        }

        // GET api/<ListaERController>/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }

        // POST api/<ListaERController>
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/<ListaERController>/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/<ListaERController>/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
