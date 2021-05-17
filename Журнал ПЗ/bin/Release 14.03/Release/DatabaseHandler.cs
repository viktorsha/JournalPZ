using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace График_ПЗ
{
    class DatabaseHandler
    {
        public static void FillDb(AppContext db, DataTable dt, List<string> cellColorList)
        {
            if (db.Employees.ToList().Count != 0)
            {
                db.Employees.RemoveRange(db.Employees.ToList());
                db.SaveChanges();
            }

            Employee employee;
            int counter = 4;
            foreach (DataRow row in dt.Rows)
            {
                employee = new Employee();
                foreach (DataColumn column in dt.Columns)
                {
                    if (Convert.ToString(column) == "№ п,п")
                    {

                        employee.OrderId = Convert.ToInt32(row[column].ToString());

                    }

                    else if (Convert.ToString(column) == "ФИО")
                        employee.Name = row[column].ToString();
                    else if (Convert.ToString(column) == "Должность, профессия")
                        employee.Position = row[column].ToString();
                    else if (Convert.ToString(column) == "Структурное подразделение")
                        employee.Department = row[column].ToString();
                    else if (Convert.ToString(column) == "Дата проверки знаний по ОТ и оказанию первой помощи (после прохождения 40-часового обучения) факт")
                    {
                        try
                        {
                            employee.ExaminationDateFact = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.ExaminationDateFact = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дата проверки знаний по ОТ и оказанию первой помощи (после прохождения 40-часового обучения) план")
                    {
                        try
                        {
                            employee.ExaminationDatePlan = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.ExaminationDatePlan = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дата проверки знаний по ОТ, ПУЭ, ПТЭ, ПожБ (согласно Правил работы с персоналом) факт")
                    {
                        try
                        {
                            employee.ExaminationComplexDateFact = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.ExaminationComplexDateFact = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дата проверки знаний по ОТ, ПУЭ, ПТЭ, ПожБ (согласно Правил работы с персоналом) план")
                    {
                        try
                        {
                            employee.ExaminationComplexDatePlan = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.ExaminationComplexDatePlan = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дата аттестации по промышленной безопасности факт")
                    {
                        try
                        {
                            employee.AttestationDateFact = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.AttestationDateFact = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дата аттестации по промышленной безопасности план")
                    {
                        try
                        {
                            employee.AttestationDatePlan = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.AttestationDatePlan = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дата прохождения пожарно-технического минимума факт")
                    {
                        try
                        {
                            employee.PbminimumPassDateFact = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.PbminimumPassDateFact = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дата прохождения пожарно-технического минимума план")
                    {
                        try
                        {
                            employee.PbminimumPassDatePlan = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.PbminimumPassDatePlan = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дата проведения медосмотра предв")
                    {
                        try
                        {
                            employee.MedicalCheckDateFact = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.MedicalCheckDateFact = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дата проведения медосмотра период")
                    {
                        try
                        {
                            employee.MedicalCheckDatePlan = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.MedicalCheckDatePlan = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Табельный номер" && row[column] != DBNull.Value)
                        employee.TabNumber = Convert.ToInt32(row[column].ToString());
                    else if (Convert.ToString(column) == "Дата рождения")
                        employee.BirthDate = row[column].ToString();
                    else if (Convert.ToString(column) == "Дата проведения вводного инструктажа")
                    {
                        try
                        {
                            employee.EntryDate = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.EntryDate = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дата перевода")
                    {
                        try
                        {
                            employee.RelocationDate = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.RelocationDate = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дата проведения первичного инструктажа")
                    {
                        try
                        {
                            employee.PrimaryInstructionDate = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.PrimaryInstructionDate = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Стажировка даты с__ по__")
                    {
                        try
                        {
                            employee.InternshipDate = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.InternshipDate = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Стажировка реквизиты приказа")
                        employee.InternshipDetails = row[column].ToString();
                    else if (Convert.ToString(column) == "Дублирование даты с__ по__")
                    {
                        try
                        {
                            employee.DublicationDate = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.DublicationDate = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Дублирование реквизиты приказа")
                        employee.DublicationDetails = row[column].ToString();
                    else if (Convert.ToString(column) == "Допуск к самостоятельной работе дата")
                    {
                        try
                        {
                            employee.IndependentDate = (DateTime.Parse(row[column].ToString()).ToString("dd.MM.yyyy"));
                        }
                        catch (Exception)
                        {
                            employee.IndependentDate = row[column].ToString();
                        }
                    }
                    else if (Convert.ToString(column) == "Допуск к самостоятельной работе реквизиты приказа")
                        employee.IndependentDetails = row[column].ToString();
                    else if (Convert.ToString(column) == "Расторгнут договор, переведен")
                        employee.ExtraStatus = row[column].ToString();
                }
                employee.CellColor = cellColorList[counter];
                counter++;
                db.Employees.Add(employee);

                db.SaveChanges();

            }

            dt.Clear();

        }
    }
}
