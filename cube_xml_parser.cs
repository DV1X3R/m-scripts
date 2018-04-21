using System;
using System.Xml.Linq;
using System.IO;

namespace CubeMetaExporter
{
    class Program
    {
        static void Main(string[] args)
        {
            string cubesPath = "C:\\"
            foreach(var cubeFile in Directory.GetFiles(cubesPath, "*.cube"))
            {
                var xmlCube = XElement.Load(cubeFile);
                string ns = "{" + xmlCube.Attribute("xmlns").Value + "}";

                string cubeName = xmlCube.Element(ns + "Name").Value;

                foreach (var measureGroup in xmlCube.Elements(ns + "MeasureGroups").Elements())
                {
                    string measureGroupName = measureGroup.Element(ns + "Name").Value;
                    string measureGroupViewName = measureGroup.Element(ns + "Measures").Element(ns + "Measure").Element(ns + "Source").Element(ns + "Source").Element(ns + "TableID").Value;

                    foreach (var dimension in measureGroup.Element(ns + "Dimensions").Elements())
                    {
                        string measureGroupDimension = dimension.Element(ns + "CubeDimensionID").Value;
                        Console.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}", cubeName, measureGroupViewName, measureGroupName, measureGroupDimension));

                    }
                }
            }
            Console.ReadLine();
        }

    }
}
