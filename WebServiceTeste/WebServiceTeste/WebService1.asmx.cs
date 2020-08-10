using System.Web.Services;
using System.Xml.Serialization;

//https://forums.asp.net/t/2090206.aspx?how+to+remove+arrayof+xml+response+in+wcf+Rest+service

namespace WebServiceTeste
{
    /// <summary>
    /// Summary description for WebService1
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]

    public class WebService1 : System.Web.Services.WebService
    {
        [WebMethod]
        public Pessoa GetPessoaXML()
        {
            Pessoa _oPessoa = new Pessoa
            {
                PessoaID = "12349227898",
                AtributoA = "AtributoAXXXXX#######",
                AtributoB = "AtributoBXXXXX#######",
                AtributoC = "AtributoCXXXXX#######"
            };

            return _oPessoa;
        }
        [WebMethod]
        //[return: XmlRoot(ElementName = "Projects")]
        //[return: XmlRoot(ElementName = "xxxxxProjects")]
        public Envelope[] GetEnvelopesXML()
        {
            Envelope[] Env = new Envelope[]
            {
                new Envelope()
                {
                    EnvelopeID = "1234",
                    EnvelopeValor = "1234.45"
                },
                new Envelope()
                {
                    EnvelopeID = "1235",
                    EnvelopeValor = "1233.78"
                }
            };

            return Env;
        }

        

        [WebMethod]
        public string HelloWorld()
        {
            return "Hello World";
        }


        [WebMethod]
        public string DigaAlgo(string Frase)
        {
            return Frase + "#### TOKEN TESTEção";
        }

        [WebMethod]
        public int Soma(int a, int b)
        {
            return a + b;
        }
    }
}
