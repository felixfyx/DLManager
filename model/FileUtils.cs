using System.IO;
using System.Xml.Serialization;

namespace DLManager.Models
{
    public static class FileUtils
    {
        public static void SaveAddressBook(AddressBook addressBook, string filePath)
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var serializer = new XmlSerializer(typeof(AddressBook));
            using (var writer = new StreamWriter(filePath))
            {
                serializer.Serialize(writer, addressBook);
            }
        }

        public static AddressBook LoadAddressBook(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return null;
            }

            try
            {
                var serializer = new XmlSerializer(typeof(AddressBook));
                using (var reader = new StreamReader(filePath))
                {
                    return (AddressBook)serializer.Deserialize(reader);
                }
            }
            catch
            {
                // Handle exceptions (e.g., log the error)
                return null;
            }
        }
    }
}
