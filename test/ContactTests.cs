using System;
using DLManager.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DLManager.Tests
{
    [TestClass]
    public class ContactTests
    {
        [TestMethod]
        public void Constructor_AssignsAllProperties()
        {
            // Act
            var contact = new Contact("John Doe", "john.doe@example.com", false);

            // Assert
            Assert.AreEqual("John Doe", contact.DisplayName);
            Assert.AreEqual("john.doe@example.com", contact.SmtpAddress);
            Assert.IsFalse(contact.IsDL);
        }
    }
}
