using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Xml;

namespace EPiServer.ConnectForSharePoint.Models.ViewModels
{
    public class ListItemViewModel
    {
        public string Name { get; set; }
        public string XmlData { get; set; }

        public List<Entry> Entrys
        {
                get
                {
                    var list = new List<Entry>();
                try
                {
                    using (XmlReader reader = XmlReader.Create(new StringReader(XmlData)))
                    {
                        bool inProperties = false;
                        bool inHyperLink = false;
                        string name = "";
                        string value = "";
                        string type = "";
                        string description = "";
                        string url = "";
                        string lastElementName = "";

                        Entry vt = new Entry();

                        // Parse the file and display each of the nodes.
                        while (reader.Read())
                        {
                            switch (reader.NodeType)
                            {
                                case XmlNodeType.Element:

                                    if (inProperties && inHyperLink != true)
                                    {
                                        name = reader.Name;
                                        type = reader.GetAttribute("m:type");
                                        if (string.Equals(type, "SP.FieldUrlValue", StringComparison.OrdinalIgnoreCase))
                                            inHyperLink = true;
                                    }

                                    if (string.Equals(reader.Name, "m:properties", StringComparison.OrdinalIgnoreCase))
                                        inProperties = true;

                                    lastElementName = reader.Name;
                                    break;
                                case XmlNodeType.Text:
                                    if (inProperties)
                                    {
                                        if (!inHyperLink)
                                            value = reader.Value;
                                        else
                                        {
                                            if (string.Equals(lastElementName, "d:Description", StringComparison.OrdinalIgnoreCase))
                                                description = reader.Value;

                                            if (string.Equals(lastElementName, "d:Url", StringComparison.OrdinalIgnoreCase))
                                                url = reader.Value;
                                        }
                                    }
                                    break;

                                case XmlNodeType.XmlDeclaration:
                                case XmlNodeType.ProcessingInstruction:
                                    Console.WriteLine(reader.Name + " " + reader.Value);
                                    break;
                                case XmlNodeType.Comment:
                                    Console.WriteLine(reader.Value);
                                    break;

                                case XmlNodeType.EndElement:
                                    if (inProperties)
                                    {
                                        if (!inHyperLink)
                                        {
                                            if (!string.IsNullOrEmpty(name))
                                            {
                                                list.Add(new Entry()
                                                {
                                                    Name = name,
                                                    Type = type,
                                                    Value = value,
                                                    Description = description,
                                                    Url = url
                                                });
                                                name = type = value = description = url = "";
                                            }
                                        }
                                        else // in hyperlink
                                        {
                                            if (string.Equals(name, reader.Name, StringComparison.OrdinalIgnoreCase))
                                            {
                                                inHyperLink = false;
                                                list.Add(new Entry()
                                                {
                                                    Name = name,
                                                    Type = type,
                                                    Value = value,
                                                    Description = description,
                                                    Url = url
                                                });
                                                name = type = value = description = url = "";
                                            }
                                        }
                                    }

                                    if (string.Equals(reader.Name, "m:properties", StringComparison.OrdinalIgnoreCase))
                                        inProperties = false;

                                    break;
                            }
                        }
                    }
                } catch
                {

                }
                return list;
            }
        }
    }

    /// <summary>
    /// Represents a single property of the list item
    /// </summary>
    public class Entry
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public string Type { get; set; }
        public string Url { get; set; }
        public string Description { get; set; }

        public override string ToString()
        {
            return "[" + Name + "," + Value + "," + Type + "," + Description + "," + Url + "]";
        }
    }
}