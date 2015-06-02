using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Mailer.Class
{
    internal class MailHelper
    {
        public bool IsValidEmail(string emailaddress)
        {
            try
            {
                MailAddress m = new MailAddress(emailaddress);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        public void sparsujIWyslij(Dictionary<string, string> slownikWartosci, string tekst)
        {
            char[] separators = new char[] { ' ','\t','\n' };
            string[] tab = tekst.Split(separators);
            string email = "";
            StringBuilder tekstDoWyslania = new StringBuilder();
            foreach (var item in tab)
            {
                if (item.Length > 0)
                {
                    if (item[0] == '$')
                    {
                        foreach (var item2 in slownikWartosci)
                        {
                            if (item2.Key == item.Substring(1, item.Length - 1))
                            {
                                tekstDoWyslania.Append(" " + item2.Value + " ");
                            }
                        }
                    }
                    else
                    {
                        tekstDoWyslania.Append(" " + item + " ");
                    }
                }               
            }
            foreach (var item in slownikWartosci)
            {
                if (IsValidEmail(item.Value))
                {
                    email += item.Value + ";";
                }
            }
            string command = "mailto:" + email + "?subject=Test&body=" + tekstDoWyslania.ToString();
            Process.Start(command);
        }
    }
}
