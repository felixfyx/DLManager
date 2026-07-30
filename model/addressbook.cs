using System.Xml.Serialization;
using System.Collections.Generic;

namespace DLManager.Models
{
    [XmlRoot("AddressBook")]
    public class AddressBook
    {
        [XmlArray("DistributionLists")]
        [XmlArrayItem("DistributionList")]
        public List<DistributionList> DistributionLists { get; set; } = new List<DistributionList>();

        /// <summary>
        /// Adds a distribution list to the address book if it doesn't already exist.
        /// </summary>
        /// <param name="distributionList">The distribution list to add.</param>
        public void AddDistributionList(DistributionList distributionList)
        {
            if (DistributionLists.Find(dl => dl.SmtpAddress == distributionList.SmtpAddress) != null)
            {
                return;
            }

            DistributionLists.Add(distributionList);
        }

        /// <summary>
        /// Removes a distribution list from the address book by its SMTP address.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address of the distribution list to remove.</param>
        public void RemoveDistributionList(string smtpAddress)
        {
            foreach (var distributionList in DistributionLists)
            {
                if (distributionList.SmtpAddress == smtpAddress)
                {
                    DistributionLists.Remove(distributionList);
                    return;
                }
            }
        }

        /// <summary>
        /// Gets a distribution list by its SMTP address.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address of the distribution list to retrieve.</param>
        /// <returns>The distribution list if found; otherwise, null.</returns>
        public DistributionList GetDistributionList(string smtpAddress)
        {
            foreach (var distributionList in DistributionLists)
            {
                if (distributionList.SmtpAddress == smtpAddress)
                {
                    return distributionList;
                }
            }
            return null;
        }
    }
}