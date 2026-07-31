using System;
using System.Collections.Generic;
using DLManager.Models;
using DLManager.Controllers;
using System.Runtime.CompilerServices;

namespace DLManager
{
    public class Program
    {

        static Contact contact1 = new Contact("John Doe", "john.doe@example.com", false);
        static Contact contact2 = new Contact("John Doe2", "john.doe2@example.com", false);
        static Contact contact3 = new Contact("John Doe3", "john.doe3@example.com", false);
        static Contact contact4 = new Contact("John Doe4", "john.doe4@example.com", false);
        static Contact contact5 = new Contact("John Doe5", "john.doe5@example.com", false);

        // 2 users in a DL
        private static void TestDL(AddressBookManager abm)
        {
            var tmpDL = new DistributionList("Team A", "team.a@example.com", DateTime.Now);
            tmpDL.AddContact(contact1);
            tmpDL.AddContact(contact2);

            var recipients = new List<Contact>()
            {
                tmpDL
            };

            abm.ProcessRecipientList(recipients);
        }

        // Test with 2 DL with 2 users in each
        private static void TestDL2(AddressBookManager abm)
        {
            var tmpDL = new DistributionList("Team A", "team.a@example.com", DateTime.Now);
            tmpDL.AddContact(contact1);
            tmpDL.AddContact(contact2);
            var tmpDL2 = new DistributionList("Team B", "team.b@example.com", DateTime.Now);
            tmpDL2.AddContact(contact3);
            tmpDL2.AddContact(contact4);

            var recipients = new List<Contact>()
            {
                tmpDL,
                tmpDL2
            };

            abm.ProcessRecipientList(recipients);
        }

        // Test with 2 DLs and 1 more user in the list
        private static void TestDL3(AddressBookManager abm)
        {
            var tmpDL = new DistributionList("Team A", "team.a@example.com", DateTime.Now);
            tmpDL.AddContact(contact1);
            tmpDL.AddContact(contact2);
            var tmpDL2 = new DistributionList("Team B", "team.b@example.com", DateTime.Now);
            tmpDL2.AddContact(contact3);
            tmpDL2.AddContact(contact4);

            var recipients = new List<Contact>()
            {
                tmpDL,
                tmpDL2,
                contact5
            };

            abm.ProcessRecipientList(recipients);
        }

        // One nested DL
        private static void TestDL4(AddressBookManager abm)
        {
            var tmpDL = new DistributionList("Team A", "team.a@example.com", DateTime.Now);
            tmpDL.AddContact(contact1);
            tmpDL.AddContact(contact2);
            var tmpDL2 = new DistributionList("Team B", "team.b@example.com", DateTime.Now);
            tmpDL2.AddContact(contact3);
            tmpDL2.AddContact(contact4);
            var tmpDL3 = new DistributionList("Team C", "team.c@example.com", DateTime.Now);
            tmpDL3.AddContact(tmpDL);

            var recipients = new List<Contact>()
            {
                tmpDL,
                tmpDL2,
                tmpDL3
            };

            abm.ProcessRecipientList(recipients);
        }

        // Duplicate DL
        private static void TestDL5(AddressBookManager abm)
        {
            var tmpDL = new DistributionList("Team A", "team.a@example.com", DateTime.Now);
            tmpDL.AddContact(contact1);
            tmpDL.AddContact(contact2);
            var tmpDL2 = new DistributionList("Team B", "team.b@example.com", DateTime.Now);
            tmpDL2.AddContact(contact3);
            tmpDL2.AddContact(contact4);
            var tmpDL3 = new DistributionList("Team C", "team.c@example.com", DateTime.Now);
            tmpDL3.AddContact(tmpDL);

            var recipients = new List<Contact>()
            {
                tmpDL,
                tmpDL2,
                tmpDL3,
                tmpDL
            };

            abm.ProcessRecipientList(recipients);
        }

        // Circular dependency
        private static void TestDL6(AddressBookManager abm)
        {
            var tmpDL = new DistributionList("Team A", "team.a@example.com", DateTime.Now);
            tmpDL.AddContact(contact1);
            tmpDL.AddContact(contact2);
            var tmpDL2 = new DistributionList("Team B", "team.b@example.com", DateTime.Now);
            tmpDL2.AddContact(contact3);
            tmpDL2.AddContact(contact4);

            tmpDL.AddContact(tmpDL2);
            tmpDL2.AddContact(tmpDL);

            var recipients = new List<Contact>()
            {
                tmpDL,
                tmpDL2
            };

            abm.ProcessRecipientList(recipients);
        }

        // Problematic case replication
        private static void TestDL7(AddressBookManager abm)
        {
            var tmpDL = new DistributionList("Team A", "team.a@example.com", DateTime.Now);
            tmpDL.AddContact(contact1);

            var tmpDL2 = new DistributionList("Team B", "team.b@example.com", DateTime.Now);
            tmpDL2.AddContact(contact2);
            tmpDL2.AddContact(contact3);
            tmpDL2.AddContact(contact4);
            tmpDL2.AddContact(contact5);

            var tmpDL3 = new DistributionList("Team C", "team.c@example.com", DateTime.Now);
            tmpDL3.AddContact(contact2);
            tmpDL3.AddContact(contact3);
            tmpDL3.AddContact(contact4);
            tmpDL3.AddContact(contact5);

            var tmpDL4 = new DistributionList("Team D", "team.d@example.com", DateTime.Now);
            tmpDL4.AddContact(contact2);
            tmpDL4.AddContact(contact3);
            tmpDL4.AddContact(contact4);
            tmpDL4.AddContact(contact5);

            tmpDL.AddContact(tmpDL2);
            tmpDL.AddContact(tmpDL3);
            tmpDL.AddContact(tmpDL4);

            var recipients = new List<Contact>()
            {
                tmpDL
            };

            abm.ProcessRecipientList(recipients);
        }

        private static void TestDL8(AddressBookManager abm)
        {
            var tmpDL = new DistributionList("Team A", "team.a@example.com", DateTime.Now);
            tmpDL.AddContact(contact1);

            var tmpDL2 = new DistributionList("Team B", "team.b@example.com", DateTime.Now);
            tmpDL2.AddContact(contact2);
            tmpDL2.AddContact(contact3);
            tmpDL2.AddContact(contact4);
            tmpDL2.AddContact(contact5);

            var tmpDL3 = new DistributionList("Team C", "team.c@example.com", DateTime.Now);
            tmpDL3.AddContact(contact2);
            tmpDL3.AddContact(contact3);
            tmpDL3.AddContact(contact4);
            tmpDL3.AddContact(contact5);

            var tmpDL4 = new DistributionList("Team D", "team.d@example.com", DateTime.Now);
            tmpDL4.AddContact(contact2);
            tmpDL4.AddContact(contact3);
            tmpDL4.AddContact(contact4);
            tmpDL4.AddContact(contact5);

            var tmpDL5 = new DistributionList("Team E", "team.e@example.com", DateTime.Now);
            tmpDL5.AddContact(tmpDL2);
            tmpDL5.AddContact(tmpDL3);

            tmpDL.AddContact(tmpDL2);
            tmpDL.AddContact(tmpDL3);
            tmpDL.AddContact(tmpDL4);
            tmpDL.AddContact(tmpDL5);

            var recipients = new List<Contact>()
            {
                tmpDL
            };

            abm.ProcessRecipientList(recipients);
        }

        private static void TestDL9(AddressBookManager abm)
        {
            var tmpDL = new DistributionList("Team A", "team.a@example.com", DateTime.Now);
            tmpDL.AddContact(contact1);

            var tmpDL2 = new DistributionList("Team B", "team.b@example.com", DateTime.Now);
            tmpDL2.AddContact(contact2);
            tmpDL2.AddContact(contact3);
            tmpDL2.AddContact(contact4);
            tmpDL2.AddContact(contact5);

            var tmpDL3 = new DistributionList("Team C", "team.c@example.com", DateTime.Now);
            tmpDL3.AddContact(contact2);
            tmpDL3.AddContact(contact3);
            tmpDL3.AddContact(contact4);
            tmpDL3.AddContact(contact5);

            var tmpDL4 = new DistributionList("Team D", "team.d@example.com", DateTime.Now);
            tmpDL4.AddContact(contact2);
            tmpDL4.AddContact(contact3);
            tmpDL4.AddContact(contact4);
            tmpDL4.AddContact(contact5);

            var tmpDL5 = new DistributionList("Team E", "team.e@example.com", DateTime.Now);

            // Double nest
            tmpDL3.AddContact(tmpDL4);

            tmpDL5.AddContact(tmpDL2);
            tmpDL5.AddContact(tmpDL3);


            tmpDL.AddContact(tmpDL2);
            tmpDL.AddContact(tmpDL3);
            tmpDL.AddContact(tmpDL4);
            tmpDL.AddContact(tmpDL5);

            var recipients = new List<Contact>()
            {
                tmpDL
            };

            abm.ProcessRecipientList(recipients);
        }

        static void Main(string[] args)
        {
            var abm = new AddressBookManager();

            TestDL9(abm);

            abm.PrintAddressBook();
        }
    }
}
