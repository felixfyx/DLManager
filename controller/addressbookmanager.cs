using System;
using System.Collections.Generic;
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
            this.expiry = expiry ?? TimeSpan.FromSeconds(30); // Default expiry of 1 minute if not provided

            // Check if we have an address book file to load
            if (File.Exists(defaultABMPath))
            {
                Console.WriteLine($"Loading address book from {defaultABMPath}");
                addressBook = FileUtils.LoadAddressBook(defaultABMPath);

                // If the address book was loaded successfully, invalidate expired distribution lists and return
                if (addressBook != null)
                {
                    InvalidateExpiredDistributionLists();
                    Console.WriteLine("Loaded address book successfully.");
                    return;
                }
            }

            // If no address book file exists / failed to load, create a new one with default values
            Console.WriteLine($"Failed to load address book from {defaultABMPath}. Creating a new address book.");
            addressBook = new AddressBook();
            Save();
        }

        /// <summary>
        /// Save address book to the default path.
        /// </summary>
        public void Save()
        {
            FileUtils.SaveAddressBook(addressBook, defaultABMPath);
        }

        private bool IsDLStale(DistributionList distributionList)
        {
            return (DateTime.Now - distributionList.CachedTime) > expiry;
        }

        // TODO: Add a function to process the whole recipient list
        public void ProcessRecipientList(List<Contact> recipients)
        {
            foreach (var recipient in recipients)
            {
                if (recipient.IsDL)
                {
                    var dl = GetDistributionList(recipient.SmtpAddress);
                    // If we didn't get any distribution list with the given SMTP address,
                    // we can say that the DL was not saved in the address book so we should create a new entry.
                    if (dl == null)
                    {
                        dl = recipient as DistributionList;
                        ProcessDistributionList(dl);
                    }
                }
                else
                {
                    Console.WriteLine($"Ignoring recipient {recipient.DisplayName} <{recipient.SmtpAddress}> as it is not a distribution list.");
                }
            }

            Save(); // Save the address book
        }

        /// <summary>
        /// Processes a distribution list and adds it to the address book if it doesn't already exist.
        /// Skips processing if the distribution list already exists in the address book.
        /// </summary>
        /// <param name="distributionList"></param>
        public void ProcessDistributionList(DistributionList distributionList)
        {
            // Check if the distribution list already exists in the address book
            var currDL = GetDistributionList(distributionList.SmtpAddress);
            Console.WriteLine($"Processing distribution list {distributionList.DisplayName} <{distributionList.SmtpAddress}>.");
            if (currDL != null)
            {
                Console.WriteLine($"Distribution list {currDL.DisplayName} <{currDL.SmtpAddress}> already exists.");
                return; // Distribution list already exists, skip processing
            }
            else
            {
                Console.WriteLine($"Adding new distribution list {distributionList.DisplayName} <{distributionList.SmtpAddress}> to the address book.");

                // Create a new DL and add entry to the address book
                currDL = new DistributionList(distributionList.DisplayName, distributionList.SmtpAddress, DateTime.Now);
                AddDistributionList(currDL); // Add the new distribution list to the address book
            }

            // Process each contact in the distribution list.
            foreach (var contact in distributionList.Contacts)
            {
                if (contact.IsDL)
                {
                    var dl = GetDistributionList(contact.SmtpAddress);
                    // If we didn't get any distribution list with the given SMTP address,
                    // we can say that the DL was not saved in the address book so we should create a new entry.
                    if (dl == null)
                    {
                        dl = contact as DistributionList;
                        Console.WriteLine($"Processing nested distribution list {dl.DisplayName} <{dl.SmtpAddress}>.");
                        ProcessDistributionList(dl);
                    }
                }

                AddContactToDistributionList(currDL, contact);
            }
        }

        public void AddContactToDistributionList(DistributionList distributionList, Contact contact)
        {
            if (contact.IsDL)
            {
                contact = new Contact(contact.DisplayName, contact.SmtpAddress, true);
            }

            distributionList.AddContact(contact);
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