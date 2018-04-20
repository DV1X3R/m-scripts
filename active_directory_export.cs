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

                        if (results.Properties["givenName"].Count != 0)
                            Output0Buffer.FirstName = results.Properties["givenName"][0].ToString();
                        if (results.Properties["sn"].Count != 0)
                            Output0Buffer.LastName = results.Properties["sn"][0].ToString();
                        if (results.Properties["sAMAccountName"].Count != 0)
                            Output0Buffer.UserLogin = results.Properties["sAMAccountName"][0].ToString();

                        Output0Buffer.UserGroupName = groupDN.Substring(3, (groupDN.IndexOf(',') - 3));
                    }
                }
            }

        }
    }

}
