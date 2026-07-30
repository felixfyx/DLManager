using System;
using DLManager.Models;
using DLManager.Controllers;

namespace DLManager
{
    public class Program
    {
        static void Main(string[] args)
        {
            var abm = new AddressBookManager();
            var contact1 = new Contact("John Doe", "john.doe@example.com", false);
            var contact2 = new Contact("John Doe2", "john.doe2@example.com", false);
            var contact3 = new Contact("John Doe3", "john.doe3@example.com", false);
            var contact4 = new Contact("John Doe4", "john.doe4@example.com", false);
            var tmpDL = new DistributionList("Team A", "team.a@example.com", DateTime.Now);
            var tmpDL2 = new DistributionList("Team B", "team.b@example.com", DateTime.Now);
            var tmpDL3 = new DistributionList("Team C", "team.c@example.com", DateTime.Now);
            tmpDL.AddContact(contact1);
            tmpDL.AddContact(contact2);
            tmpDL2.AddContact(contact3);
            tmpDL2.AddContact(contact4);
            tmpDL3.AddContact(tmpDL); // Add tmpDL to tmpDL3 to test nested DLs

            // Try add tmpDL to itself to prevent duplicate entries
            tmpDL.AddContact(tmpDL);

            abm.AddDistributionList(tmpDL3);
            abm.Save();
            abm.PrintAddressBook();
        }
    }
}
