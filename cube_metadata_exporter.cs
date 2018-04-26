using Microsoft.AnalysisServices;
using System.Collections.Generic;

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
                        Output0Buffer.ObjectType = "Cube";
                        Output0Buffer.ObjectName = OLAPCube.Name;
                        Output0Buffer.MeasureGroupName = OLAPMeasureGroup.Name;
                        Output0Buffer.MeasureGroupDimensionName = OLAPMeasureGroupDimension.CubeDimension.Name;
                        Output0Buffer.DimensionName = OLAPMeasureGroupDimension.CubeDimension.DimensionID;
                    }
                }

                foreach (Perspective OLAPPerspective in OLAPCube.Perspectives)
                {
                    var perspectiveDimensions = new List<CubeDimension>();
                    var mappedDimensions = new List<CubeDimension>();

                    foreach (PerspectiveDimension OLAPPerspectiveDimension in OLAPPerspective.Dimensions)
                    {
                        perspectiveDimensions.Add(OLAPPerspectiveDimension.CubeDimension);
                    }

                    foreach (PerspectiveMeasureGroup OLAPPerspectiveMeasureGroup in OLAPPerspective.MeasureGroups)
                    {
                        foreach (MeasureGroupDimension OLAPMeasureGroupDimension in OLAPPerspectiveMeasureGroup.MeasureGroup.Dimensions)
                        {
                            if (perspectiveDimensions.Exists((x) => x.Equals(OLAPMeasureGroupDimension.CubeDimension)))
                            {
                                mappedDimensions.Add(OLAPMeasureGroupDimension.CubeDimension);

                                Output0Buffer.AddRow();
                                Output0Buffer.DatabaseName = OLAPDatabase.Name;
                                Output0Buffer.ObjectType = "Perspective";
                                Output0Buffer.ObjectName = OLAPPerspective.Name;
                                Output0Buffer.MeasureGroupName = OLAPPerspectiveMeasureGroup.MeasureGroup.Name;
                                Output0Buffer.MeasureGroupDimensionName = OLAPMeasureGroupDimension.CubeDimension.Name;
                                Output0Buffer.DimensionName = OLAPMeasureGroupDimension.CubeDimension.DimensionID;
                            }
                        }
                    }

                    perspectiveDimensions.RemoveAll((x) => mappedDimensions.Contains(x));

                    foreach(CubeDimension UnmappedDimension in perspectiveDimensions)
                    {
                        Output0Buffer.AddRow();
                        Output0Buffer.DatabaseName = OLAPDatabase.Name;
                        Output0Buffer.ObjectType = "Perspective";
                        Output0Buffer.ObjectName = OLAPPerspective.Name;
                        Output0Buffer.MeasureGroupName = null;
                        Output0Buffer.MeasureGroupDimensionName = UnmappedDimension.Name;
                        Output0Buffer.DimensionName = UnmappedDimension.DimensionID;
                    }

                }
            }
        }
    }

}
