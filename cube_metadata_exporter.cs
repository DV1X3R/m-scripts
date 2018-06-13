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
                // CUBES
                foreach (MeasureGroup OLAPMeasureGroup in OLAPCube.MeasureGroups)
                {
                    foreach (Measure OLAPMeasure in OLAPMeasureGroup.Measures)
                    {
                        Output0Buffer.AddRow();
                        Output0Buffer.DatabaseName = OLAPDatabase.Name;
                        Output0Buffer.SourceType = "Cube";
                        Output0Buffer.SourceName = OLAPCube.Name;
                        Output0Buffer.MeasureGroupName = OLAPMeasureGroup.Name;
                        Output0Buffer.ObjectType = "Measure";
                        Output0Buffer.ObjectName = OLAPMeasure.Name;
                    }

                    foreach (MeasureGroupDimension OLAPMeasureGroupDimension in OLAPMeasureGroup.Dimensions)
                    {
                        Output0Buffer.AddRow();
                        Output0Buffer.DatabaseName = OLAPDatabase.Name;
                        Output0Buffer.SourceType = "Cube";
                        Output0Buffer.SourceName = OLAPCube.Name;
                        Output0Buffer.MeasureGroupName = OLAPMeasureGroup.Name;
                        Output0Buffer.ObjectType = "Dimension";
                        if (OLAPMeasureGroupDimension.CubeDimension.Name == OLAPMeasureGroupDimension.CubeDimension.DimensionID)
                            Output0Buffer.ObjectName = OLAPMeasureGroupDimension.CubeDimension.Name;
                        else
                            Output0Buffer.ObjectName = string.Format("{0} ({1})"
                                , OLAPMeasureGroupDimension.CubeDimension.Name
                                , OLAPMeasureGroupDimension.CubeDimension.DimensionID);
                    }
                }

                
                // PERSPECTIVES
                foreach (Perspective OLAPPerspective in OLAPCube.Perspectives)
                {
                    var mappedDimensions = new List<CubeDimension>();
                    var perspectiveDimensions = new List<CubeDimension>();
                    foreach (PerspectiveDimension OLAPPerspectiveDimension in OLAPPerspective.Dimensions)
                        perspectiveDimensions.Add(OLAPPerspectiveDimension.CubeDimension);

                    foreach (PerspectiveMeasureGroup OLAPPerspectiveMeasureGroup in OLAPPerspective.MeasureGroups)
                    {
                        foreach (PerspectiveMeasure OLAPPerspectiveMeasure in OLAPPerspectiveMeasureGroup.Measures)
                        {
                            Output0Buffer.AddRow();
                            Output0Buffer.DatabaseName = OLAPDatabase.Name;
                            Output0Buffer.SourceType = "Perspective";
                            Output0Buffer.SourceName = OLAPPerspective.Name;
                            Output0Buffer.MeasureGroupName = OLAPPerspectiveMeasureGroup.MeasureGroup.Name;
                            Output0Buffer.ObjectType = "Measure";
                            Output0Buffer.ObjectName = OLAPPerspectiveMeasure.Measure.Name;
                        }

                        foreach (MeasureGroupDimension OLAPMeasureGroupDimension in OLAPPerspectiveMeasureGroup.MeasureGroup.Dimensions)
                        {
                            if (perspectiveDimensions.Exists((x) => x.Equals(OLAPMeasureGroupDimension.CubeDimension)))
                            {
                                mappedDimensions.Add(OLAPMeasureGroupDimension.CubeDimension);

                                Output0Buffer.AddRow();
                                Output0Buffer.DatabaseName = OLAPDatabase.Name;
                                Output0Buffer.SourceType = "Perspective";
                                Output0Buffer.SourceName = OLAPPerspective.Name;
                                Output0Buffer.MeasureGroupName = OLAPPerspectiveMeasureGroup.MeasureGroup.Name;
                                Output0Buffer.ObjectType = "Dimension";
                                if (OLAPMeasureGroupDimension.CubeDimension.Name == OLAPMeasureGroupDimension.CubeDimension.DimensionID)
                                    Output0Buffer.ObjectName = OLAPMeasureGroupDimension.CubeDimension.Name;
                                else
                                    Output0Buffer.ObjectName = string.Format("{0} ({1})"
                                        , OLAPMeasureGroupDimension.CubeDimension.Name
                                        , OLAPMeasureGroupDimension.CubeDimension.DimensionID);
                            }
                        }
                    }

                    // PERSPECTIVE UNMAPPED DIMENSIONS

                    perspectiveDimensions.RemoveAll((x) => mappedDimensions.Contains(x));

                    foreach (CubeDimension UnmappedDimension in perspectiveDimensions)
                    {
                        Output0Buffer.AddRow();
                        Output0Buffer.DatabaseName = OLAPDatabase.Name;
                        Output0Buffer.SourceType = "Perspective";
                        Output0Buffer.SourceName = OLAPPerspective.Name;
                        Output0Buffer.MeasureGroupName = null;
                        Output0Buffer.ObjectType = "Dimension";
                        if (UnmappedDimension.Name == UnmappedDimension.DimensionID)
                            Output0Buffer.ObjectName = UnmappedDimension.Name;
                        else
                            Output0Buffer.ObjectName = string.Format("{0} ({1})"
                                , UnmappedDimension.Name
                                , UnmappedDimension.DimensionID);
                    }

                }
                
            }
        }
    }

}
