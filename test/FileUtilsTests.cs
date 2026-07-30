using System;
using System.IO;
using DLManager.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DLManager.Tests
{
    [TestClass]
    public class FileUtilsTests
    {
        private string tempFilePath;

        [TestInitialize]
        public void TestInitialize()
        {
            tempFilePath = Path.Combine(Path.GetTempPath(), $"addressbook-{Guid.NewGuid()}.xml");
        }

        [TestCleanup]
        public void TestCleanup()
        {
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
        }

        [TestMethod]
        public void SaveAddressBook_CreatesFileOnDisk()
        {
            // Arrange
            var addressBook = new AddressBook();

            // Act
            FileUtils.SaveAddressBook(addressBook, tempFilePath);

            // Assert
            Assert.IsTrue(File.Exists(tempFilePath));
        }

        [TestMethod]
        public void SaveAddressBook_ThenLoadAddressBook_RoundTripsData()
        {
            // Arrange
            var addressBook = new AddressBook();
            var distributionList = new DistributionList("Team", "team@example.com", DateTime.Now);
            distributionList.AddContact(new Contact("John Doe", "john.doe@example.com", false));
            addressBook.AddDistributionList(distributionList);

            // Act
            FileUtils.SaveAddressBook(addressBook, tempFilePath);
            var loaded = FileUtils.LoadAddressBook(tempFilePath);

            // Assert
            Assert.IsNotNull(loaded);
            Assert.AreEqual(1, loaded.DistributionLists.Count);
            Assert.AreEqual("team@example.com", loaded.DistributionLists[0].SmtpAddress);
            Assert.AreEqual(1, loaded.DistributionLists[0].Contacts.Count);
            Assert.AreEqual("john.doe@example.com", loaded.DistributionLists[0].Contacts[0].SmtpAddress);
        }

        [TestMethod]
        public void LoadAddressBook_ReturnsNull_WhenFileDoesNotExist()
        {
            // Act
            var loaded = FileUtils.LoadAddressBook(tempFilePath);

            // Assert
            Assert.IsNull(loaded);
        }
    }
}
