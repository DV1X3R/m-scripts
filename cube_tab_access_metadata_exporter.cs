using Microsoft.AnalysisServices.Tabular;
using System.Text.RegularExpressions;

[Microsoft.SqlServer.Dts.Pipeline.SSISScriptComponentEntryPointAttribute]
public class ScriptMain : UserComponent
{
    public override void CreateNewOutputRows()
    {
        Server server = new Server();
        server.Connect(Variables.ASTabularServerCRConnectionString);
        Regex rx = new Regex(@"_.*_*-"); // Exclude Temp DB ConnexysReporting_UserName_3f2d309e-237d-4755-8e89-69b33f5f5992

        foreach (Role role in server.Roles)
        {
            foreach (Microsoft.AnalysisServices.RoleMember member in role.Members)
            {
                Output0Buffer.AddRow();
                Output0Buffer.DatabaseType = "Tabular";
                Output0Buffer.DatabaseName = server.Name;
                Output0Buffer.SourceType = "Server";
                Output0Buffer.SourceName = server.Name;
                Output0Buffer.RoleName = role.Name;
                Output0Buffer.RoleMemberName = member.Name;
            }
        }

        foreach (Database database in server.Databases)
        {
            if (rx.Matches(database.Name).Count != 0) continue;

            foreach (ModelRole role in database.Model.Roles)
            {
                foreach (ModelRoleMember member in role.Members)
                {
                    Output0Buffer.AddRow();
                    Output0Buffer.DatabaseType = "Tabular";
                    Output0Buffer.DatabaseName = server.Name;
                    Output0Buffer.SourceType = "Model";
                    Output0Buffer.SourceName = database.Name;
                    Output0Buffer.RoleName = role.Name;
                    Output0Buffer.RoleMemberName = member.Name;

                    var permission = role.ModelPermission.ToString();
                    Output0Buffer.Read = permission == "Read" || permission == "ReadRefresh" || permission == "Administrator" ? "Allowed" : "None";
                    Output0Buffer.Process = permission == "ReadRefresh" || permission == "Administrator" ? "True" : "False";
                    Output0Buffer.Administrator = permission == "Administrator" ? "True" : "False";
                }
            }
        }
    }

}
