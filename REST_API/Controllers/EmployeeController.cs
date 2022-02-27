using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using График_ПЗ;

namespace REST_API.Controllers
{
    public class EmployeeController : ApiController
    {
        [HttpGet]
        public IHttpActionResult Get(int id)
        {
            try
            {
                using (EmployeeDBEntities entities = new EmployeeDBEntities())
                {
                    var emp = entities.Employees.FirstOrDefault(em => em.ID == id);
                    if (emp != null)
                    {
                        return Ok(emp);
                    }
                    else
                    {
                        return Content(HttpStatusCode.NotFound, "Employee with Id: " + id + " not found");
                    }
                }
            }
            catch (Exception ex)
            {
                return Content(HttpStatusCode.BadRequest, ex);

            }
        }

        [HttpGet]
        public IHttpActionResult GetList()
        {
            try
            {
                using (EmployeeDBEntities entities = new EmployeeDBEntities())
                {
                    var emp = entities.Employees.ToList();
                    if (emp != null)
                    {
                        return Ok(emp);
                    }
                    else
                    {
                        return Content(HttpStatusCode.NotFound, "Employee list is empty");
                    }
                }
            }
            catch (Exception ex)
            {
                return Content(HttpStatusCode.BadRequest, ex);

            }
        }

        [HttpGet]
        public IHttpActionResult FormLateList()
        {
            try
            {
                using (EmployeeDBEntities entities = new EmployeeDBEntities())
                {
                    var emp = entities.Employees(em => DateTime.Compare(DateTime.Parse(em.ExaminationDatePlan), DateTime.Now) < 0);
                    if (emp != null)
                    {
                        return Ok(emp);
                    }
                    else
                    {
                        return Content(HttpStatusCode.NotFound, "No late examinations");
                    }
                }
            }
            catch (Exception ex)
            {
                return Content(HttpStatusCode.BadRequest, ex);

            }
        }

        [HttpGet]
        public IHttpActionResult FormRangeList(DateTime from, DateTime to)
        {
            try
            {
                using (EmployeeDBEntities entities = new EmployeeDBEntities())
                {
                    var emp = entities.Employees(em => (DateTime.Compare(DateTime.Parse(em.ExaminationDatePlan), DateTime.Parse(from)) > 0) && (DateTime.Compare(DateTime.Parse(em.ExaminationDatePlan), DateTime.Parse(to)) < 0));
                    if (emp != null)
                    {
                        return Ok(emp);
                    }
                    else
                    {
                        return Content(HttpStatusCode.NotFound, "No examinations in chosen range");
                    }
                }
            }
            catch (Exception ex)
            {
                return Content(HttpStatusCode.BadRequest, ex);

            }
        }

        [HttpPost]
        public HttpResponseMessage AddEmployee([FromBody] Employee employee)
        {
            try
            {
                using (EmployeeDBEntities entities = new EmployeeDBEntities())
                {
                    entities.Employees.Add(employee);
                    entities.SaveChanges();
                    var res = Request.CreateResponse(HttpStatusCode.Created, employee);
                    res.Headers.Location = new Uri(Request.RequestUri + employee.Id.ToString());
                    return res;
                }
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex);
            }
        }

        [HttpPut]
        public HttpResponseMessage UpdateEmployee(int id, [FromBody] Employee emp)
        {
            try
            {
                using (EmployeeDBEntities entities = new EmployeeDBEntities())
                {
                    var employee = entities.Employees.Where(em => em.ID == id).FirstOrDefault();
                    if (employee != null)
                    {
                        if (!string.IsNullOrWhiteSpace(emp.Name))
                            employee.FIO = emp.Name;

                        if (!string.IsNullOrWhiteSpace(emp.Department))
                            employee.Department = emp.Department;

                        if (!string.IsNullOrWhiteSpace(emp.Position))
                            employee.Position = emp.Position;

                        if (!string.IsNullOrWhiteSpace(emp.ExaminationDateFact))
                            employee.ExaminationDateFact = emp.ExaminationDateFact;

                        if (!string.IsNullOrWhiteSpace(emp.ExaminationDatePlan))
                            employee.ExaminationDatePlan = emp.ExaminationDatePlan;

                        if (!string.IsNullOrWhiteSpace(emp.ExaminationComplexDateFact))
                            employee.ExaminationComplexDateFact = emp.ExaminationComplexDateFact;

                        if (!string.IsNullOrWhiteSpace(emp.ExaminationComplexDatePlan))
                            employee.ExaminationComplexDatePlan = emp.ExaminationComplexDatePlan;

                        if (!string.IsNullOrWhiteSpace(emp.AttestationDateFact))
                            employee.AttestationDateFact = emp.AttestationDateFact;

                        if (!string.IsNullOrWhiteSpace(emp.AttestationDatePlan))
                            employee.AttestationDatePlan = emp.AttestationDatePlan;

                        if (!string.IsNullOrWhiteSpace(emp.PbminimumPassDateFact))
                            employee.PbminimumPassDateFact = emp.PbminimumPassDateFact;

                        if (!string.IsNullOrWhiteSpace(emp.PbminimumPassDatePlan))
                            employee.PbminimumPassDatePlan = emp.PbminimumPassDatePlan;

                        if (!string.IsNullOrWhiteSpace(emp.MedicalCheckDateFact))
                            employee.MedicalCheckDateFact = emp.MedicalCheckDateFact;

                        if (!string.IsNullOrWhiteSpace(emp.MedicalCheckDatePlan))
                            employee.MedicalCheckDatePlan = emp.MedicalCheckDatePlan;

                        if (!string.IsNullOrWhiteSpace(emp.TabNumber))
                            employee.TabNumber = emp.TabNumber;

                        if (!string.IsNullOrWhiteSpace(emp.BirthDate))
                            employee.BirthDate = emp.BirthDate;

                        if (!string.IsNullOrWhiteSpace(emp.EntryDate))
                            employee.EntryDate = emp.EntryDate;


                        entities.SaveChanges();
                        var res = Request.CreateResponse(HttpStatusCode.OK, "Employee with id" + id + " updated");
                        res.Headers.Location = new Uri(Request.RequestUri + id.ToString());
                        return res;
                    }
                    else
                    {
                        return Request.CreateErrorResponse(HttpStatusCode.NotFound, "Employee with id" + id + " is not found!");
                    }
                }
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex);
            }
        }

        [HttpDelete]
        public HttpResponseMessage DeleteEmployee(int id)
        {
            try
            {
                using (EmployeeDBEntities entities = new EmployeeDBEntities())
                {
                    var employee = entities.Employees.Where(emp => emp.ID == id).FirstOrDefault();
                    if (employee != null)
                    {
                        entities.Employees.Remove(employee);
                        entities.SaveChanges();
                        return Request.CreateResponse(HttpStatusCode.OK, "Employee with id" + id + " Deleted");
                    }
                    else
                    {
                        return Request.CreateErrorResponse(HttpStatusCode.NotFound, "Employee with id" + id + " is not found!");
                    }
                }
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex);
            }
        }
    }
}
