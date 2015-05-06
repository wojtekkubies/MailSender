using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSenderApp.Class
{
    public class Szablony
    {
        public string Nazwa { get; set; }
        public string Tekst { get; set; }
    }

    public class RootObject
    {
        public List<Szablony> szablony { get; set; }
    }
}
