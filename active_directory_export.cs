using System;
using System.DirectoryServices;
using System.Collections.Generic;

[Microsoft.SqlServer.Dts.Pipeline.SSISScriptComponentEntryPointAttribute]
public class ScriptMain : UserComponent
{
    public override void CreateNewOutputRows()
    {
        using (DirectorySearcher ds = new DirectorySearcher()) //"LDAP://"
        {
            var userGroups = new List<string>() { "*" };
            var userGroupsDN = new List<string>();

            ds.SearchScope = SearchScope.Subtree;
            ds.PageSize = 100000; // This will page through the records 1000 at a time

            // Select all the DistinguishedNames for selected groups 
            foreach (string group in userGroups)
            {
                ds.Filter = "(&(objectCategory=Group)(cn=" + group + "))";
                using (SearchResultCollection src = ds.FindAll())
                    foreach (SearchResult results in src)
                        userGroupsDN.Add(results.Properties["distinguishedName"][0].ToString());
            }

            foreach (string groupDN in userGroupsDN)
            {
                ds.Filter = "(&(objectCategory=User)(memberOf=" + groupDN + "))";
                using (SearchResultCollection src = ds.FindAll())
                {
                    foreach (SearchResult sr in src)
                    {
                        Output0Buffer.AddRow();
                        Output0Buffer.FirstName = GetStringProperty(sr, "givenName");
                        Output0Buffer.LastName = GetStringProperty(sr, "sn");
                        string userLogin = GetStringProperty(sr, "sAMAccountName");
                        Output0Buffer.UserLogin = userLogin;
                        Output0Buffer.UserMail = GetStringProperty(sr, "mail");
                        Output0Buffer.UserDomainName = GetStringProperty(sr, "userPrincipalName");
                        Output0Buffer.UserAdPath = GetStringProperty(sr, "distinguishedName");
                        Output0Buffer.UserGroupName = groupDN.Substring(3, (groupDN.IndexOf(',') - 3));
                        Output0Buffer.UserGroupAdPath = groupDN;

                        Output0Buffer.Company = GetStringProperty(sr, "company");
                        Output0Buffer.Description = GetStringProperty(sr, "description");
                        Output0Buffer.DisplayName = GetStringProperty(sr, "displayName");
                        Output0Buffer.EmployeeNr = GetStringProperty(sr, "employeeNumber");
                        Output0Buffer.EmployeeID = GetStringProperty(sr, "employeeID");

                        Output0Buffer.LastLogon = GetFileTimeProperty(sr, "lastLogon");
                        Output0Buffer.AccountExpires = GetFileTimeProperty(sr, "accountExpires");
                        Output0Buffer.WhenCreated = GetStringProperty(sr, "whenCreated");
                        Output0Buffer.UserAccountControl = IsActiveAccount(sr);

                        Output0Buffer.Department = GetStringProperty(sr, "department");
                        Output0Buffer.Manager = GetStringProperty(sr, "manager");
                        Output0Buffer.Mobile = GetStringProperty(sr, "mobile");
                        Output0Buffer.Title = GetStringProperty(sr, "title");

                        Output0Buffer.ObjectGUID = GetObjectGUID(sr);
                        Output0Buffer.IsCiberEmployee = GetStringProperty(sr, "distinguishedName").IndexOf(",OU=Ciber,") > 0 ? true : false;

                    }
                }
            }

        }
    }

    private string GetStringProperty(SearchResult sr, string propertyName)
        => sr.Properties[propertyName].Count != 0 ?
        sr.Properties[propertyName][0].ToString() : null;

    private string GetFileTimeProperty(SearchResult sr, string propertyName)
        => (sr.Properties[propertyName].Count != 0 && (Int64)sr.Properties[propertyName][0] != 0x7FFFFFFFFFFFFFFF) ?
        DateTime.FromFileTime(
            (Int64)sr.Properties[propertyName][0])
            .ToString()
        : null;

    private string IsActiveAccount(SearchResult sr)
        => sr.Properties["userAccountControl"].Count != 0 ?
        (!Convert.ToBoolean(
            (int)sr.Properties["userAccountControl"][0] & 0x0002))
            .ToString()
        : null;

    private string GetObjectGUID(SearchResult sr)
    {
        if (sr.Properties["objectGUID"].Count != 0)
        {
            byte[] binaryData = sr.Properties["objectGUID"][0] as byte[];
            string strHex = BitConverter.ToString(binaryData);
            Guid id = new Guid(strHex.Replace("-", ""));
            return id.ToString().Replace("-", "");
        }
        else return null;
    }

    private bool IsGroupMember(string groupDistinguishedName, string sAMAccountName)
    {
        DirectorySearcher ds = new DirectorySearcher();
        ds.Filter = string.Format("(&(memberOf:1.2.840.113556.1.4.1941:={0})(objectCategory=person)(objectClass=user)(sAMAccountName={1}))", groupDistinguishedName, sAMAccountName);
        SearchResult src = ds.FindOne();
        return src != null;
    }

}
