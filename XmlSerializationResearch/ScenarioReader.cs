using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Reflection;
using System.Xml.Serialization;
using System.IO;

namespace XmlSerializationResearch
{
    public class ScenarioReader
    {
        Dictionary<string, Type> transformCommandTypes = new Dictionary<string, Type>();

        public ScenarioReader()
        {
            Assembly commandAssembly = Assembly.GetAssembly(typeof(ITransformCommand));
            Type[] types = commandAssembly.GetExportedTypes();

            foreach (Type type in types)
            {
                Type iTransformCommand = type.GetInterface("ITransformCommand");
                if (iTransformCommand != null)
                    transformCommandTypes.Add(type.Name, type);
            }
        }

        public IList<ITransformCommand> LoadCommands(string scenarioPath)
        {
            List<ITransformCommand> commands = new List<ITransformCommand>();
            XmlDocument doc = new XmlDocument();
            doc.Load(scenarioPath);

            // элементы верхнего уровня представляют собой отдельные команды
            foreach (XmlNode node in doc.DocumentElement.ChildNodes)
            {
                if (node.Name != "#comment")
                {
                    XmlSerializer serializer = new XmlSerializer(transformCommandTypes[node.Name]);
                    ITransformCommand command = (ITransformCommand)serializer.Deserialize(new MemoryStream(Encoding.UTF8.GetBytes(node.OuterXml)));
                    commands.Add(command);
                }
            }

            return commands;
        }
    }
}
