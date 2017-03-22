using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateCustomXMLPart
{
    class Program
    {
        static void Main(string[] args)
        {
            CMDocument SSLTestDoc = new CMDocument();
            SSLTestDoc.UpdateSSLContactInfo();
        }
    }
}
