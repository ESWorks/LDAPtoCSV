using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LDAPtoCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            // create your domain context
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, args[0]);

            // create your principal searcher   
            PrincipalSearcher srch = new PrincipalSearcher(new UserPrincipal(ctx));

            // find all matches
            foreach (var found in srch.FindAll())
            {
                if (found.GetType() == typeof(UserPrincipal))
                {
                    UserPrincipal user = (UserPrincipal)found;
                    string first = user.SamAccountName;
                    string last = user.Surname;
                    string email = user.EmailAddress;
                    string org = user.VoiceTelephoneNumber;
                    if (string.IsNullOrEmpty(email)) continue;
                    Console.WriteLine($@"{first},{last},{email},{org}");
                }      
            }
        }

        //old method to read GAL
        private static void Read()
        {
            List<string> emails = GetGAL(args[1],args[2], args[0]);
            emails.Sort();

            foreach (string item in emails)
            {
                Console.WriteLine(item);
            }
        }

        // search by directory, required username and pass.
        public static List<string> GetGAL(string UserName, string Password, string server)
        {
            try
            {
                List<string> ReturnArray = new List<string>();
                DirectoryEntry deDirEntry = new DirectoryEntry(server,
                                                                    UserName,
                                                                    Password,
                                                                    AuthenticationTypes.Secure);
                DirectorySearcher mySearcher = new DirectorySearcher(deDirEntry);
               
                string sFilter = String.Format("(&(objectCategory=person)(objectClass=user))");

                mySearcher.Filter = sFilter;
                mySearcher.PropertiesToLoad.AddRange(new[] { "mail","sn","GivenName","DisplayName" });
                mySearcher.Sort.Direction = System.DirectoryServices.SortDirection.Ascending;
                mySearcher.PageSize = 1000;
                
                SearchResultCollection results;
                results = mySearcher.FindAll();
                
                foreach (SearchResult resEnt in results)
                {
                    string sn = "";
                    string mail = "";
                    string gv = "";
                    string dn = "";
                    ResultPropertyCollection propcoll = resEnt.Properties;
                    foreach (string key in propcoll.PropertyNames)
                    {
                        if (key == "mail")
                        {
                            foreach (object values in propcoll[key])
                            {
                                mail = values.ToString();
                            }
                        }
                        if (key == "sn")
                        {
                            foreach (object values in propcoll[key])
                            {
                                sn = values.ToString();
                            }
                        }
                        if (key == "GivenName")
                        {
                            foreach (object values in propcoll[key])
                            {
                                gv = values.ToString();
                            }
                        }
                        if (key == "DisplayName")
                        {
                            foreach (object values in propcoll[key])
                            {
                                dn = values.ToString();
                            }
                        }
                    }
                    ReturnArray.Add(dn+","+sn+","+","+mail+","+gv);
                }
                return ReturnArray;
            }
            catch
            {
                return null;
            }
        }
    }
}
