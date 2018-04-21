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
            foreach (Cube OLAPCube in OLAPDatabase.Cubes)
            {
                foreach (MeasureGroup OLAPMeasureGroup in OLAPCube.MeasureGroups)
                {
                    foreach (MeasureGroupDimension OLAPMeasureGroupDimension in OLAPMeasureGroup.Dimensions)
                    {
                        Output0Buffer.AddRow();
                        Output0Buffer.DatabaseName = OLAPDatabase.Name;
                        Output0Buffer.CubeName = OLAPCube.Name;
                        Output0Buffer.MeasureGroupName = OLAPMeasureGroup.Name;
                        Output0Buffer.MeasureGroupDimensionName = OLAPMeasureGroupDimension.CubeDimension.Name;
                        Output0Buffer.DimensionName = OLAPMeasureGroupDimension.CubeDimension.DimensionID;
                    }
                }
            }
        }
    }
}
