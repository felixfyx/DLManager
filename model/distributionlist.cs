using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace DLManager.Models
{
    public class DistributionList : Contact
    {
        [XmlArray("Contacts")]
        [XmlArrayItem("Contact")]
        public List<Contact> Contacts { get; set; } = new List<Contact>();

        [XmlElement("CachedTime")]
        public DateTime CachedTime { get; set; }

        public DistributionList()
        {
        }

        public DistributionList(string displayName, string smtpAddress, DateTime cachedTime)
            : base(displayName, smtpAddress, true)
        {
            CachedTime = cachedTime;
        }

        public void AddContact(Contact contact)
        {
            if (Contacts.Find(c => c.SmtpAddress == contact.SmtpAddress) != null ||
                contact.SmtpAddress == this.SmtpAddress)
            {
                return;
            }

            /*
            // Flatten DL and then process it at the to level of the address book.
            if (contact.IsDL)
            {
                contact = new Contact(contact.DisplayName, contact.SmtpAddress, true);
            }
            */

            Contacts.Add(contact);
        }
        
        public void RemoveContact(string smtpAddress)
        {
            foreach (var contact in Contacts)
            {
                if (contact.SmtpAddress == smtpAddress)
                {
                    Contacts.Remove(contact);
                    return;
                }
            }
        }
    }
}