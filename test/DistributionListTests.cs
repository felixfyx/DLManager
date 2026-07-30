using System;
using DLManager.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DLManager.Tests
{
    [TestClass]
    public class DistributionListTests
    {
        [TestMethod]
        public void AddContact_AddsNewContact()
        {
            // Arrange
            var list = new DistributionList("Team", "team@example.com", DateTime.Now);
            var contact = new Contact("John Doe", "john.doe@example.com", false);

            // Act
            list.AddContact(contact);

            // Assert
            Assert.AreEqual(1, list.Contacts.Count);
            Assert.AreSame(contact, list.Contacts[0]);
        }

        [TestMethod]
        public void AddContact_IgnoresDuplicateSmtpAddress()
        {
            // Arrange
            var list = new DistributionList("Team", "team@example.com", DateTime.Now);
            var original = new Contact("John Doe", "john.doe@example.com", false);
            var duplicate = new Contact("Jonathan Doe", "john.doe@example.com", false);
            list.AddContact(original);

            // Act
            list.AddContact(duplicate);

            // Assert
            Assert.AreEqual(1, list.Contacts.Count);
            Assert.AreSame(original, list.Contacts[0]);
        }

        [TestMethod]
        public void RemoveContact_RemovesMatchingContact()
        {
            // Arrange
            var list = new DistributionList("Team", "team@example.com", DateTime.Now);
            var contact = new Contact("John Doe", "john.doe@example.com", false);
            list.AddContact(contact);

            // Act
            list.RemoveContact("john.doe@example.com");

            // Assert
            Assert.AreEqual(0, list.Contacts.Count);
        }

        [TestMethod]
        public void RemoveContact_NoOpWhenAddressNotFound()
        {
            // Arrange
            var list = new DistributionList("Team", "team@example.com", DateTime.Now);
            var contact = new Contact("John Doe", "john.doe@example.com", false);
            list.AddContact(contact);

            // Act
            list.RemoveContact("someone.else@example.com");

            // Assert
            Assert.AreEqual(1, list.Contacts.Count);
        }
    }
}
