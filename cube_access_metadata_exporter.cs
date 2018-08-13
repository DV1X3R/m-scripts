using Microsoft.AnalysisServices;

[Microsoft.SqlServer.Dts.Pipeline.SSISScriptComponentEntryPointAttribute]
public class ScriptMain : UserComponent
{
    public override void CreateNewOutputRows()
    {
        Server OLAPServer = new Server();
        OLAPServer.Connect(Variables.ASCubeServerCRConnectionString);
        foreach (Database OLAPDatabase in OLAPServer.Databases)
        {
            foreach (DatabasePermission OLAPDatabasePermission in OLAPDatabase.DatabasePermissions)
            {
                foreach (RoleMember OLAPRoleMember in OLAPDatabasePermission.Role.Members)
                {
                    Output0Buffer.AddRow();
                    Output0Buffer.DatabaseName = OLAPDatabase.Name;
                    Output0Buffer.SourceType = "Database";
                    Output0Buffer.SourceName = OLAPDatabase.Name;
                    Output0Buffer.RoleName = OLAPDatabasePermission.Role.Name;
                    Output0Buffer.RoleMemberName = OLAPRoleMember.Name;
                    Output0Buffer.Read = OLAPDatabasePermission.Read.ToString();
                    Output0Buffer.Write = OLAPDatabasePermission.Write.ToString();
                    Output0Buffer.Process = OLAPDatabasePermission.Process.ToString();
                    Output0Buffer.Administrator = OLAPDatabasePermission.Administer.ToString();
                }
            }

            foreach (Cube OLAPCube in OLAPDatabase.Cubes)
            {
                foreach (CubePermission OLAPCubePermission in OLAPCube.CubePermissions)
                {
                    foreach (RoleMember OLAPRoleMember in OLAPCubePermission.Role.Members)
                    {
                        Output0Buffer.AddRow();
                        Output0Buffer.DatabaseName = OLAPDatabase.Name;
                        Output0Buffer.SourceType = "Cube";
                        Output0Buffer.SourceName = OLAPCube.Name;
                        Output0Buffer.RoleName = OLAPCubePermission.Role.Name;
                        Output0Buffer.RoleMemberName = OLAPRoleMember.Name;
                        Output0Buffer.Read = OLAPCubePermission.Read.ToString();
                        Output0Buffer.Write = OLAPCubePermission.Write.ToString();
                        Output0Buffer.Process = OLAPCubePermission.Process.ToString();
                    }
                }
            }
        }
    }

}
