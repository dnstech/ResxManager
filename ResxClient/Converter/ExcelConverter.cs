using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.Reflection;
using System.Globalization;
using System.Linq;
using ResourceManager.Converter.Exceptions;
using ResourceManager.Core;

namespace ResourceManager.Converter
{
    public class ExcelConverter : ConverterBase, IConverter
    {
        private const string urnschemasmicrosoftcomofficespreadsheet = "urn:schemas-microsoft-com:office:spreadsheet";
        private const string urnschemasmicrosoftcomofficeexcel = "urn:schemas-microsoft-com:office:excel";

        private int expandedColumnCount = 0;
        private int expandedRowCount = 0;

        public ExcelConverter(VSSolution solution) : base(solution)
        {
        }
        public ExcelConverter(VSProject project) : base(project)
        {
        }

        public XmlDocument Export()
        {
            throw new NotImplementedException();
        }

        public int Import(string filename)
        {
            int count = 0;

            XmlDocument xml = new XmlDocument();
            xml.Load(filename);

            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(xml.NameTable);
            namespaceManager.AddNamespace("", urnschemasmicrosoftcomofficespreadsheet);
            namespaceManager.AddNamespace("o", "urn:schemas-microsoft-com:office:office");
            namespaceManager.AddNamespace("x", urnschemasmicrosoftcomofficeexcel);
            namespaceManager.AddNamespace("ss", urnschemasmicrosoftcomofficespreadsheet);
            namespaceManager.AddNamespace("html", "http://www.w3.org/TR/REC-html40");

            List<VSCulture> cultures = new List<VSCulture>();

            XmlNodeList worksheets = xml.SelectNodes("/ss:Workbook/ss:Worksheet", namespaceManager);

            foreach (XmlNode worksheet in worksheets)
            {
                string projectName = worksheet.SelectSingleNode("@ss:Name", namespaceManager).Value;
                if(!Solution.Projects.ContainsKey(projectName))
                    throw new ProjectUnknownException(projectName);

                VSProject project = Solution.Projects[projectName];

                XmlNodeList nodes = worksheet.SelectNodes("ss:Table/ss:Row", namespaceManager);
                foreach (XmlNode rowNode in nodes)
                {
                    if (rowNode == rowNode.ParentNode.SelectSingleNode("ss:Row", namespaceManager))
                    {
                        XmlNodeList cellNodes = rowNode.SelectNodes("ss:Cell", namespaceManager);
                        for (int i = 2; i < cellNodes.Count; i++)
                        {
                            cultures.Add(new VSCulture(CultureInfo.GetCultureInfo(cellNodes[i].FirstChild.InnerText)));
                        }
                    }
                    else
                    {
                        string key = rowNode.ChildNodes[1].FirstChild.InnerText;
                        string id = rowNode.FirstChild.FirstChild.InnerText;

                        ResourceDataGroupBase dataGroup = null;
                        if (!project.ResxGroups[id].AllData.ContainsKey(key))
                        {
                            dataGroup = project.ResxGroups[id].CreateDataGroup(key);
                            project.ResxGroups[id].AllData.Add(key, dataGroup);
                        }
                        else
                            dataGroup = project.ResxGroups[id].AllData[key];

                        for (int i = 0; i < cultures.Count; i++)
                        {
                            XmlNode valueNode = rowNode.SelectSingleNode("ss:Cell[@ss:Index = '" + (i + 3) + "']/ss:Data", namespaceManager);
                            if (valueNode == null)
                                valueNode = rowNode.SelectSingleNode("ss:Cell[count(@ss:Index) = 0][" + (i + 3) + "]/ss:Data", namespaceManager);

                            if (valueNode != null)
                            {
                                if (!dataGroup.ResxData.ContainsKey(cultures[i].Culture))
                                {
                                    project.ResxGroups[id].SetResourceData(key, valueNode.InnerText, cultures[i].Culture);                                   
                                }
                                else
                                {
                                    dataGroup.ResxData[cultures[i].Culture].Value = valueNode.InnerText;
                                }

                                count++;
                            }
                        }
                    }
                }
            }

            return count;
        }       
    }
}
