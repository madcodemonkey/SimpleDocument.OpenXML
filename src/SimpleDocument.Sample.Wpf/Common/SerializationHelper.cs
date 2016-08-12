using System.IO;
using System.Xml.Serialization;

namespace WordOpenXMLExample1
{
    public class SerializationHelper
    {
        /// <summary>Serializes the given object to an XML string.</summary>
        public static string SerializeToXml(object objectToSerialize)
        {
            var serializer = new XmlSerializer(objectToSerialize.GetType());

            using (var memoryStream = new MemoryStream())
            {
                serializer.Serialize(memoryStream, objectToSerialize);
                memoryStream.Seek(0, SeekOrigin.Begin);

                using (var streamReader = new StreamReader(memoryStream))
                    return streamReader.ReadToEnd();
            }
        }

        public static void SerializeToXmlFile(object objectToSerialize, string filePath)
        {
            var serializer = new XmlSerializer(objectToSerialize.GetType());

            using (FileStream fileStream = File.OpenWrite(filePath))
                serializer.Serialize(fileStream, objectToSerialize);
        }
    }
}