using System;
using System.IO;
using DLManager.Models;

namespace DLManager.Controllers
{
    public sealed partial class AddressBookManager
    {
        private string defaultABMPath;
        private AddressBook addressBook;
        private TimeSpan expiry;

        public AddressBookManager(TimeSpan? expiry = null)
        {
            // Hardcoded default ABM path for loading
            defaultABMPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "addressbook.xml");
            this.expiry = expiry ?? TimeSpan.FromMinutes(1); // Default expiry of 1 minute if not provided

            // Check if we have an address book file to load
            if (File.Exists(defaultABMPath))
            {
                Console.WriteLine($"Loading address book from {defaultABMPath}");
                addressBook = FileUtils.LoadAddressBook(defaultABMPath);
            }
            else
            {
                // If no address book file exists, create a new one with default values
                Console.WriteLine($"No address book found at {defaultABMPath}. Creating a new one.");
                addressBook = new AddressBook();
                FileUtils.SaveAddressBook(addressBook, defaultABMPath);
            }

            InvalidateExpiredDistributionLists();
        }

        /*
        TODOS: 
        - Is Contact Stale
        - Invalidate expired contacts
        */

        /// <summary>
        /// Save address book to the default path.
        /// </summary>
        public void Save()
        {
            FileUtils.SaveAddressBook(addressBook, defaultABMPath);
        }

        private bool IsDLStale(DistributionList distributionList)
        {
            try
            {
                return (DateTime.Now - distributionList.CachedTime) > expiry;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking if distribution list is stale: {ex.Message}");
                return true; // If there's an error, consider it stale
            }
        }

        // TODO: Add a function to process the whole recipient list
        public void ProcessRecipientList()
        {
            throw new NotImplementedException("ProcessRecipientList is not implemented yet.");
        }

        /// <summary>
        /// Add distribution list if it does not already exist in the address book. If it exists, do not add it again.
        /// </summary>
        /// <param name="distributionList"></param>
        public void AddDistributionList(DistributionList distributionList)
        {
            // Checks already happen within this function
            addressBook.AddDistributionList(distributionList);
        }

        public void RemoveDistributionList(DistributionList distributionList)
        {
            addressBook.RemoveDistributionList(distributionList.SmtpAddress);
        }

        /// <summary>
        /// Gets a distribution list by its SMTP address.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address of the distribution list to retrieve.</param>
        /// <returns>The distribution list if found, otherwise null.</returns>
        public DistributionList GetDistributionList(string smtpAddress)
        {
            return addressBook.GetDistributionList(smtpAddress);
        }

        /// <summary>
        /// Invalidates expired distribution lists within the address book.
        /// </summary>
        public void InvalidateExpiredDistributionLists()
        {
            addressBook.DistributionLists.RemoveAll(dl => IsDLStale(dl));
        }

        /// <summary>
        /// Debug helper that prints the entire address book to the console.
        /// </summary>
        public void PrintAddressBook()
        {
            Console.WriteLine($"AddressBook: {addressBook.DistributionLists.Count} distribution list(s)");

            foreach (var dl in addressBook.DistributionLists)
            {
                Console.WriteLine($"- {dl.DisplayName} <{dl.SmtpAddress}> (CachedTime: {dl.CachedTime}, Stale: {IsDLStale(dl)})");

                foreach (var contact in dl.Contacts)
                {
                    Console.WriteLine($"    - {contact.DisplayName} <{contact.SmtpAddress}> (IsDL: {contact.IsDL})");
                }
            }
        }
    }
}