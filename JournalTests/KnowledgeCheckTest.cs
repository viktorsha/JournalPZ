using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using REST_API.Controllers;
using System;
using System.Collections.Generic;
using System.Net.Http;
using График_ПЗ;

namespace JournalTests
{
    [TestClass]
    public class KnowledgeCheckTest
    {
        private readonly EmployeeController _controller;

        public KnowledgeCheckTest()
        {
            _controller = new EmployeeController();
        }

        [TestMethod]
        public void LateKnowledgeCheckFormedCorrectly()
        {
            var mockContext = new Mock<EmployeeContext>();
            Employee employee = new Employee
            {
                Name = "jajaja",
                ExaminationDatePlan = DateTime.Parse("25.02.2022").ToString()
            };
            mockContext.Setup(m => m.Employees).Returns(employee);
            var checkList = _controller.FormLateList();
            Assert.AreEqual(employee, checkList);
        }

        [TestMethod]
        public void LateKnowledgeCheckInRangeFormedCorrectly()
        {
            var mockContext = new Mock<EmployeeContext>();
            Employee employee = new Employee
            {
                Name = "jajaja",
                ExaminationDatePlan = DateTime.Parse("25.02.2022").ToString()
            };
            mockContext.Setup(m => m.Employees).Returns(employee);
            var checkList = _controller.FormRangeList();
            Assert.AreEqual(employee, checkList);
        }

        [TestMethod]
        public void AddingEmployeesReturnsSuccessStatusCode()
        {
            var mockContext = new Mock<EmployeeContext>();
            Employee employee = new Employee
            {
                Name = "jajaja",
                ExaminationDatePlan = DateTime.Parse("25.02.2022").ToString()
            };
            HttpResponseMessage response = _controller.AddEmployee(employee);

            Assert.AreEqual(System.Net.HttpStatusCode.OK, response.StatusCode);
        }

        [TestMethod]
        public void DeletingEmployeesReturnsCorrectList()
        {
            var mockContext = new Mock<EmployeeContext>();
            Employee employee = new Employee
            {
                Name = "jajaja",
                ExaminationDatePlan = DateTime.Parse("25.02.2022").ToString()
            };
            mockContext.Setup(emp => emp.Employees).Returns(employee);
            HttpResponseMessage response = _controller.DeleteEmployee(0);

            Assert.AreEqual(System.Net.HttpStatusCode.OK, response.StatusCode);
        }

        [TestMethod]
        public void GetEmployeeReturnsCorrectRecord()
        {
            var mockContext = new Mock<EmployeeContext>();
            Employee employee = new Employee
            {
                Name = "jajaja",
                ExaminationDatePlan = DateTime.Parse("25.02.2022").ToString()
            };
            mockContext.Setup(emp => emp.Employees).Returns(employee);
            HttpResponseMessage response = _controller.Get(0);
            Assert.AreEqual(System.Net.HttpStatusCode.OK, response.StatusCode);
        }

        [TestMethod]
        public void GetEmployeeListReturnsCorrectRecord()
        {
            var mockContext = new Mock<EmployeeContext>();
            var employeeList = new List<Employee>()
            {
                new Employee
                {
                    Name = "jajaja",
                    ExaminationDatePlan = DateTime.Parse("25.02.2022").ToString()
                },
                new Employee
                {
                    Name = "hello",
                    ExaminationDatePlan = DateTime.Parse("20.05.2021").ToString()
                }
            };
            mockContext.Setup(emp => emp.Employees).Returns(employeeList);
            HttpResponseMessage response = _controller.GetList();
            Assert.AreEqual(System.Net.HttpStatusCode.OK, response.StatusCode);
        }

        [TestMethod]
        public void PutEmployeeListUpdatesTheRecord()
        {
            var mockContext = new Mock<EmployeeContext>();
            var employeeList = new List<Employee>()
            {
                new Employee
                {
                    Name = "jajaja",
                    ExaminationDatePlan = DateTime.Parse("25.02.2022").ToString()
                },
                new Employee
                {
                    Name = "hello",
                    ExaminationDatePlan = DateTime.Parse("20.05.2021").ToString()
                }
            };
            Employee employee = new Employee
            {
                Name = "FIO",
                ExaminationDatePlan = DateTime.Parse("15.02.2019").ToString()
            };
            mockContext.Setup(emp => emp.Employees).Returns(employeeList);
            HttpResponseMessage response = _controller.UpdateEmployee(1, employee);
            Assert.AreEqual(System.Net.HttpStatusCode.OK, response.StatusCode);
        }
    }
}
