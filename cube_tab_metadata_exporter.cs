using Microsoft.AnalysisServices.Tabular;
using System.Collections.Generic;
using System.Text.RegularExpressions;

public class Relation
{
    public readonly string Fact;
    public readonly string Dimension;

    public Relation(string fact, string dimension)
    {
        Fact = fact;
        Dimension = dimension;
    }
}

[Microsoft.SqlServer.Dts.Pipeline.SSISScriptComponentEntryPointAttribute]
public class ScriptMain : UserComponent
{
    public override void CreateNewOutputRows()
    {
        Server server = new Server();
        server.Connect(Variables.ASTabularServerCRConnectionString);
        Regex rx = new Regex(@"_.*_*-"); // Exclude Temp DB ConnexysReporting_UserName_GUID

        foreach (Database database in server.Databases)
        {

            if (rx.Matches(database.Name).Count != 0) continue;

            var model = database.Model;
            var relationships = new List<Relation>();
            foreach (Relationship relationship in model.Relationships)
                relationships.Add(new Relation(relationship.FromTable.Name, relationship.ToTable.Name));

            // Dimensions
            foreach (Relation relation in relationships)
            {
                Output0Buffer.AddRow();
                Output0Buffer.DatabaseType = "Tabular";
                Output0Buffer.DatabaseName = database.Name;
                Output0Buffer.SourceType = "Model";
                Output0Buffer.SourceName = model.Name;
                Output0Buffer.MeasureGroupName = relation.Fact;
                Output0Buffer.ObjectType = "Dimension";
                Output0Buffer.ObjectName = relation.Dimension;
            }

            foreach (Table table in model.Tables)
            {
                // Measures
                if (table.Measures.Count > 0)
                {
                    foreach (Measure measure in table.Measures)
                    {
                        Output0Buffer.AddRow();
                        Output0Buffer.DatabaseType = "Tabular";
                        Output0Buffer.DatabaseName = database.Name;
                        Output0Buffer.SourceType = "Model";
                        Output0Buffer.SourceName = model.Name;
                        Output0Buffer.MeasureGroupName = table.Name;
                        Output0Buffer.ObjectType = "Measure";
                        Output0Buffer.ObjectName = measure.Name;
                    }
                }
                // Unmapped tables
                else if (!relationships.Exists(x => x.Dimension.Equals(table.Name)))
                {
                    Output0Buffer.AddRow();
                    Output0Buffer.DatabaseType = "Tabular";
                    Output0Buffer.DatabaseName = database.Name;
                    Output0Buffer.SourceType = "Model";
                    Output0Buffer.SourceName = model.Name;
                    Output0Buffer.MeasureGroupName = null;
                    Output0Buffer.ObjectType = "Table";
                    Output0Buffer.ObjectName = table.Name;
                }
            }

            // Perspectives
            foreach (Perspective perspective in model.Perspectives)
            {
                var perspectiveFacts = new List<string>();
                foreach (PerspectiveTable table in perspective.PerspectiveTables)
                {
                    if (table.PerspectiveMeasures.Count > 0)
                        perspectiveFacts.Add(table.Name);
                }

                foreach (PerspectiveTable table in perspective.PerspectiveTables)
                {
                    if (relationships.Exists(x => x.Dimension.Equals(table.Name)))
                    {
                        // Dimension
                        foreach (Relation relation in relationships.FindAll(x => x.Dimension.Equals(table.Name)))
                        {
                            if (perspectiveFacts.Exists(x => x.Equals(relation.Fact)))
                            {
                                Output0Buffer.AddRow();
                                Output0Buffer.DatabaseType = "Tabular";
                                Output0Buffer.DatabaseName = database.Name;
                                Output0Buffer.SourceType = "Perspective";
                                Output0Buffer.SourceName = perspective.Name;
                                Output0Buffer.MeasureGroupName = relation.Fact;
                                Output0Buffer.ObjectType = "Dimension";
                                Output0Buffer.ObjectName = relation.Dimension;
                            }
                        }
                    }
                    else if (table.PerspectiveMeasures.Count > 0)
                    {
                        // Fact Measures
                        foreach (PerspectiveMeasure measure in table.PerspectiveMeasures)
                        {
                            Output0Buffer.AddRow();
                            Output0Buffer.DatabaseType = "Tabular";
                            Output0Buffer.DatabaseName = database.Name;
                            Output0Buffer.SourceType = "Perspective";
                            Output0Buffer.SourceName = perspective.Name;
                            Output0Buffer.MeasureGroupName = table.Name;
                            Output0Buffer.ObjectType = "Measure";
                            Output0Buffer.ObjectName = measure.Name;
                        }
                    }
                    else
                    {
                        // Unmapped table
                        Output0Buffer.AddRow();
                        Output0Buffer.DatabaseType = "Tabular";
                        Output0Buffer.DatabaseName = database.Name;
                        Output0Buffer.SourceType = "Perspective";
                        Output0Buffer.SourceName = perspective.Name;
                        Output0Buffer.MeasureGroupName = null;
                        Output0Buffer.ObjectType = "Table";
                        Output0Buffer.ObjectName = table.Name;
                    }
                }
            }

        }
    }

}
