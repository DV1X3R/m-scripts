using System.DirectoryServices;
using System.Collections.Generic;

[Microsoft.SqlServer.Dts.Pipeline.SSISScriptComponentEntryPointAttribute]
public class ScriptMain : UserComponent
{
    public override void CreateNewOutputRows()
    {
        //using (DirectorySearcher ds = new DirectorySearcher("LDAP://"))
        using (DirectorySearcher ds = new DirectorySearcher()
        {
            var userGroups = new List<string>() { "*" };
            var userGroupsDN = new List<string>();

            ds.SearchScope = SearchScope.Subtree;
            ds.PageSize = 5000; // This will page through the records 1000 at a time

            // Select all the DistinguishedNames of selected groups 
            foreach (string groupS in userGroups)
            {
                ds.Filter = "(&(objectCategory=Group)(cn=" + groupS + "))";
                using (SearchResultCollection src = ds.FindAll())
                    foreach (SearchResult results in src)
                        userGroupsDN.Add(results.Properties["DistinguishedName"][0].ToString());
            }

            foreach (string groupDN in userGroupsDN)
            {
                ds.Filter = "(&(objectCategory=User)(memberOf=" + groupDN + "))";
                using (SearchResultCollection src = ds.FindAll())
                {
                    foreach (SearchResult results in src)
                    {
                        Output0Buffer.AddRow();

                        Output0Buffer.FirstName = GetProperty(sr, "givenName");
                        Output0Buffer.LastName = GetProperty(sr, "sn");
                        Output0Buffer.UserLogin = GetProperty(sr, "sAMAccountName");
                        Output0Buffer.UserMail = GetProperty(sr, "mail");
                        Output0Buffer.UserAdPath = GetProperty(sr, "DistinguishedName");
                        Output0Buffer.UserDomainName = GetProperty(sr, "UserPrincipalName");
                        Output0Buffer.UserGroupName = groupDN.Substring(3, (groupDN.IndexOf(',') - 3));
                        Output0Buffer.UserGroupAdPath = groupDN;
                        
                    }
                }
            }

        }
    }
    
    public string GetProperty(SearchResult sr, string propertyName)
    {
        if (sr.Properties[propertyName].Count != 0)
            return sr.Properties[propertyName][0].ToString();
        else return null;
    }

}
