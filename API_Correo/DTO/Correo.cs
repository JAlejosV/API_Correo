using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace API_Correo.DTO
{
    public class Correo
    {
        public StructMail CorreoEmisor { get; set; }
        public List<StructMail> CorreosPara { get; set; }
        public List<StructMail> CorreosCC { get; set; }
        public List<StructMail> CorreosCCO { get; set; }
        public string Asunto { get; set; }
        public string Cuerpo { get; set; }
        public List<Adjunto> Adjuntos { get; set; }
        public Correo()
        {
            CorreoEmisor = new StructMail() { Mail = string.Empty, NameMail = string.Empty, Password = string.Empty };
            CorreosPara = new List<StructMail>();
            CorreosCC = new List<StructMail>();
            CorreosCCO = new List<StructMail>();
            Asunto = string.Empty;
            Cuerpo = string.Empty;
            Adjuntos = new List<Adjunto>();
        }
    }

    public class StructMail
    {
        public string Mail { get; set; }
        public string Password { get; set; }
        public string NameMail { get; set; }
    }

    public struct Adjunto
    {
        public byte[] archivo { get; set; }
        public string nombreArchivo { get; set; }
    }
}
