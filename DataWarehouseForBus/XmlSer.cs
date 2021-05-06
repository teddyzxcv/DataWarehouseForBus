using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.IO;
namespace DataWarehouseForBus
{
    class XmlSer
    {
        public static void SerializationToXml(List<Cloth> dtb)
        {
            FileStream fs = new FileStream("database.xml", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            XmlSerializer serializer = new XmlSerializer(typeof(List<Cloth>));
            serializer.Serialize(fs, dtb);
            fs.Close();
        }

        public static List<Cloth> DeserializationFromXml()
        {
            FileStream fs = new FileStream("database.xml", FileMode.Open, FileAccess.Read);
            XmlSerializer serializer = new XmlSerializer(typeof(List<Cloth>));
            var output = (List<Cloth>)serializer.Deserialize(fs);
            fs.Close();
            return output;
        }
    }
}
