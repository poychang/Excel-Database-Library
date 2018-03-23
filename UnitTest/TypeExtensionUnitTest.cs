using System;
using ExcelDatabaseLibrary.DataAnnotations;
using ExcelDatabaseLibrary.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTest
{
    [TestClass]
    public class TypeExtensionUnitTest
    {
        [TestMethod]
        public void GetAttribute_���w���O���ݩ����Ҽҫ�_���^���O���ݩ����Ҽҫ�����()
        {
            // Arrange: ��l�ƥؼЪ���B�̪ۨ���B��k�ѼơB�w�����G�A�άO�w���P�̪ۨ��󪺤��ʤ覡

            // Act: �I�s�ؼЪ��󪺤�k
            var tableAttribute = typeof(TableModel).GetAttribute<TableAttribute>();

            // Assert: ���ҬO�_�ŦX�w��
            Assert.AreEqual(tableAttribute.Name, "�ڪ���ƪ�");
        }
    }

    [Table(Name = "�ڪ���ƪ�")]
    internal class TableModel
    {
        [Column(Name = "�ѧO�X")]
        public string Id { get; set; }
        [Column(Name = "�إߤ��")]
        public DateTimeOffset CreateDate { get; set; }
    }
}
