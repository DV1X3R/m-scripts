using System;
using System.ServiceModel;

namespace WSDL
{
    class Program
    {
        static void Main(string[] args)
        {
            var wsdl = "?WSDL";
            var token = "";
            var connectorId = "";

            var binding = new BasicHttpBinding(BasicHttpSecurityMode.Transport) { MaxReceivedMessageSize = 2097152, MaxBufferSize = 2097152 };

            var soap = new AfasService.ConnectorAppGetSoapClient(binding, new EndpointAddress(wsdl));
            var resp = soap.GetData(token, connectorId, "", -1, -1);
            Console.WriteLine(resp);
            Console.ReadLine();
        }
    }
}
