using System;
using System.Xml.Serialization;

namespace DLManager.Models
{
    public class Contact
    {
        [XmlElement("DisplayName")]
        public string DisplayName { get; set; }
        [XmlElement("SmtpAddress")]
        public string SmtpAddress { get; set; }
        [XmlElement("IsDL")]
        public bool IsDL { get; set; }

        public Contact()
        {
        }

        public Contact(string displayName, string smtpAddress, bool isDL)
        {
            DisplayName = displayName;
            SmtpAddress = smtpAddress;
            IsDL = isDL;
        }
    }
}