using System;
using System.Data;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Data.SQLite;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;
using DataTable = System.Data.DataTable;
using Workbook = Spire.Xls.Workbook;
using Worksheet = Spire.Xls.Worksheet;
using System.Windows.Media;
using System.Windows.Controls.Primitives;
using Spire.Xls;

using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ComponentModel;

namespace График_ПЗ
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        AppContext db, db1;
        DataTable dt;
        int chosen_tab;
        public Microsoft.Office.Interop.Excel.Application APP = null;
        public Microsoft.Office.Interop.Excel.Workbook WB = null;
        public Microsoft.Office.Interop.Excel.Worksheet WS = null;
        public Range Range = null;
        public List<string> cellColorList = new List<string>();
        public List<bool> isColorChanged = new List<bool>();
        int clickedContextMenu = 0;
        public Stack<Employee> operationsDoneDetails = new Stack<Employee>();
        public Stack<string> operationsDone = new Stack<string>();
        public Stack<Employee> operationsUndoneDetails = new Stack<Employee>();
        public Stack<string> operationsUndone = new Stack<string>();
        public Dictionary<int[], string> cellColorDictionary = new Dictionary<int[], string>();
        int printing = 0;
        int[] selectedCell = new int[2];
        bool lateUsed = false;
        bool rangeUsed = false;
        public MainWindow()
        {
            try
            {
                db = new AppContext();
                db.Database.Initialize(true);
                db.Database.CreateIfNotExists();
                IEnumerable<string> tempList = from b in db.Employees.ToList()
                                               orderby b.OrderId
                                               select b.CellColor;
                cellColorList = tempList.ToList();
                for (int i = 0; i < cellColorList.Count; i++)
                    isColorChanged.Add(false);
                for (int i = 0; i<cellColorList.Count; i++)
                {
                    if (cellColorList[i] == "" )
                    {
                        cellColorList[i] = "#FFFFFFFF";
                        string com = "UPDATE Employees SET CellColor='#FFFFFFFF' WHERE OrderId='"+(i+1)+"'";
                        db.Database.ExecuteSqlCommand(com);
                        db.Dispose();
                        db = new AppContext();
                    }
                    if (cellColorList[i]!="#FFFFFFFF")
                    {
                        isColorChanged[i] = true;
                    }
                    else
                    {
                        isColorChanged[i] = false;
                    }
                }
                InitializeComponent();
                chosen_tab = 1;
                ChangeTabColor();
                DisplayInfoFull();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + " " + e.InnerException + " " + e.StackTrace);
            }
        }
        private void BtnLoadFile_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Загрузка и обработка excel файла может занять до нескольких минут. Можно сворачивать программу, пока идет загрузка. Ожидайте.", "Загрузка файла", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            switch (result)
            {
                case MessageBoxResult.OK:
                    string filePath = "";
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Filter = "Excel files (*.xls)|*.xlsx|All files (*.*)|*.*";
                    if (openFileDialog.ShowDialog() == true)
                    {
                        filePath = openFileDialog.FileName;
                        MessageBox.Show("Загрузка файла завершена успешно, ожидайте добавления данных в базу", "Загрузка файла", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    Workbook workbook = new Workbook();
                    try
                    {
                        workbook.LoadFromFile(filePath);
                        Worksheet sheet = workbook.Worksheets[0];
                        dt = sheet.ExportDataTable();
                        if (cellColorList.Count!=0)
                            cellColorList.Clear();
                        int number = sheet.Rows.Count();
                        foreach (var item in sheet.Rows)
                        {
                            var color = item.Style.Color;
                            Color color1 = Color.FromRgb(color.R, color.G, color.B);
                            if (color1.ToString().Equals("#FF000000"))
                            {
                                color1 = Color.FromRgb(255, 255, 255);
                            }
                            cellColorList.Add(color1.ToString());
                        }
                        dt.Rows.RemoveAt(0);
                        EmployeeDataGrid.ItemsSource = null;
                        EmployeeDataGrid.Items.Refresh();
                        try
                        {
                            DatabaseHandler.FillDb(db, dt, cellColorList);
                            var employeeList1 = from b in db.Employees.ToList()
                                                select b.CellColor;
                            cellColorList = employeeList1.ToList();
                            DisplayInfoFull();
                            MessageBox.Show("Данные успешно загружены!", "Загрузка файла", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        catch (Exception e1)
                        {
                            MessageBox.Show("Ошибка обработки файла! Столбцы могли быть в неправильном формате", "Ошибка обработки", MessageBoxButton.OK, MessageBoxImage.Error);
                            MessageBox.Show(e1.Message + " " + e1.InnerException + " " + e1.StackTrace);
                        }
                    }
                    catch (Exception e2)
                    {
                        MessageBox.Show("Отмена загрузки", "Статус загрузки", MessageBoxButton.OK, MessageBoxImage.Information);
                        MessageBox.Show(e2.Message + " " + e2.InnerException + " " + e2.StackTrace);
                    }
                    break;
            }
        }

        private void Tabitem1_Clicked(object sender, MouseButtonEventArgs e)
        {
            chosen_tab = 1;
            ChangeTabColor();
            DisplayInfoFull();
        }
        private void HideColumns()
        {
            examComplFactCol.Visibility = Visibility.Hidden;
            examComplPlanCol.Visibility = Visibility.Hidden;
            attestFact.Visibility = Visibility.Hidden;
            attestPlan.Visibility = Visibility.Hidden;
            pbMinFactCol.Visibility = Visibility.Hidden;
            pbMinPlanCol.Visibility = Visibility.Hidden;
            medPlanCol.Visibility = Visibility.Hidden;
            medFactCol.Visibility = Visibility.Hidden;
            tabCol.Visibility = Visibility.Hidden;
            birthCol.Visibility = Visibility.Hidden;
            entryCol.Visibility = Visibility.Hidden;
            relocCol.Visibility = Visibility.Hidden;
            primCol.Visibility = Visibility.Hidden;
        }
        private void ShowColumns()
        {
            examComplFactCol.Visibility = Visibility.Visible;
            examComplPlanCol.Visibility = Visibility.Visible;
            attestFact.Visibility = Visibility.Visible;
            attestPlan.Visibility = Visibility.Visible;
            pbMinFactCol.Visibility = Visibility.Visible;
            pbMinPlanCol.Visibility = Visibility.Visible;
            medPlanCol.Visibility = Visibility.Visible;
            medFactCol.Visibility = Visibility.Visible;
            tabCol.Visibility = Visibility.Visible;
            birthCol.Visibility = Visibility.Visible;
            entryCol.Visibility = Visibility.Visible;
            relocCol.Visibility = Visibility.Visible;
            primCol.Visibility = Visibility.Visible;
        }

        public void DisplayInfoFull()
        {
            
            EmployeeDataGrid.AutoGenerateColumns = false;
            var employeeList1 = from b in db.Employees.ToList()
                                orderby b.OrderId
                                select b;
            List<Employee> list = employeeList1.ToList();
            dt = ConvertListToDataTable(list, 26);
            EmployeeDataGrid.ItemsSource = employeeList1.ToList();
            
        }
        public void ShowInfo()
        {
            switch(chosen_tab)
            {
                case 1:
                    DisplayInfoFull();
                    break;
                case 2:
                    ExaminationShortGenerate();
                    break;
                case 4:
                    Tab4Chosen();
                    break;
                case 5:
                    Tab5Chosen();
                    break;
                case 6:
                    Tab6Chosen();
                    break;
                case 7:
                    Tab7Chosen();
                    break;
                case 8:
                    Tab8Chosen();
                    break;
                case 9:
                    Tab9Chosen();
                    break;
                case 10:
                    Tab10Chosen();
                    break;
                case 11:
                    Tab11Chosen();
                    break;
            }
        }
        public void GoBack_Click(object sender, RoutedEventArgs e)
        {
            if (operationsDoneDetails.Count!=0 && operationsDone.Count!=0)
            {
                Employee employee = operationsDoneDetails.Pop();
                string operations = operationsDone.Pop();
                string com;

                if (operations=="insert")
                {
                    cellColorList.RemoveAt(employee.OrderId - 1);
                    com = "DELETE FROM Employees WHERE OrderId='" + (employee.OrderId) + "'";
                    db.Database.ExecuteSqlCommand(com);
                    com = "UPDATE Employees SET OrderId=OrderId-1 WHERE OrderId>=" + employee.OrderId;
                    db.Database.ExecuteSqlCommand(com);
                    db.Dispose();
                    db = new AppContext();
                    ShowInfo();
                }
                else if (operations=="delete")
                {
                    cellColorList.Insert(employee.OrderId - 1, "#FFFFFFFF");
                    com = "UPDATE Employees SET OrderId=OrderId+1 WHERE OrderId>=" + employee.OrderId;
                    db.Database.ExecuteSqlCommand(com);

                    com = "INSERT INTO Employees (OrderId, TabNumber, CellColor) VALUES (" + employee.OrderId + ",0, '#FFFFFFFF')";
                    db.Database.ExecuteSqlCommand(com);

                    com = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', ExaminationDateFact='" + employee.ExaminationDateFact + "', ExaminationDatePlan='" +
                            employee.ExaminationDatePlan + "', ExaminationComplexDateFact='" + employee.ExaminationComplexDateFact +
                            "', ExaminationComplexDatePlan='" + employee.ExaminationComplexDatePlan + "', AttestationDateFact='" + employee.AttestationDateFact +
                            "', AttestationDatePlan='" + employee.AttestationDatePlan + "', PbminimumPassDateFact='" + employee.PbminimumPassDateFact +
                            "', PbminimumPassDatePlan='" + employee.PbminimumPassDatePlan + "', MedicalCheckDateFact='" + employee.MedicalCheckDateFact +
                            "', MedicalCheckDatePlan='" + employee.MedicalCheckDatePlan + "', TabNumber='" + employee.TabNumber + "', BirthDate='" + employee.BirthDate +
                            "', EntryDate='" + employee.EntryDate + "', RelocationDate='" + employee.RelocationDate + "', PrimaryInstructionDate='" +
                            employee.PrimaryInstructionDate + "', InternshipDate='" + employee.InternshipDate + "', InternshipDetails='" + employee.InternshipDetails + "'," +
                            " DublicationDate='" + employee.DublicationDate + "', DublicationDetails='" + employee.DublicationDetails + "', IndependentDate='" +
                            employee.IndependentDate + "', IndependentDetails='" + employee.IndependentDetails + "', ExtraStatus='" + employee.ExtraStatus + "', CellColor='"+employee.CellColor+"' WHERE OrderId='" + employee.OrderId + "'";
                    db.Database.ExecuteSqlCommand(com);
                    db.Dispose();
                    db = new AppContext();
                    ShowInfo();
                }
                else if (operations== "update")
                {
                    com = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', ExaminationDateFact='" + employee.ExaminationDateFact + "', ExaminationDatePlan='" +
                            employee.ExaminationDatePlan + "', ExaminationComplexDateFact='" + employee.ExaminationComplexDateFact +
                            "', ExaminationComplexDatePlan='" + employee.ExaminationComplexDatePlan + "', AttestationDateFact='" + employee.AttestationDateFact +
                            "', AttestationDatePlan='" + employee.AttestationDatePlan + "', PbminimumPassDateFact='" + employee.PbminimumPassDateFact +
                            "', PbminimumPassDatePlan='" + employee.PbminimumPassDatePlan + "', MedicalCheckDateFact='" + employee.MedicalCheckDateFact +
                            "', MedicalCheckDatePlan='" + employee.MedicalCheckDatePlan + "', TabNumber='" + employee.TabNumber + "', BirthDate='" + employee.BirthDate +
                            "', EntryDate='" + employee.EntryDate + "', RelocationDate='" + employee.RelocationDate + "', PrimaryInstructionDate='" +
                            employee.PrimaryInstructionDate + "', InternshipDate='" + employee.InternshipDate + "', InternshipDetails='" + employee.InternshipDetails + "'," +
                            " DublicationDate='" + employee.DublicationDate + "', DublicationDetails='" + employee.DublicationDetails + "', IndependentDate='" +
                            employee.IndependentDate + "', IndependentDetails='" + employee.IndependentDetails + "', ExtraStatus='" + employee.ExtraStatus + "', CellColor='"+employee.CellColor+"' WHERE OrderId='" + employee.OrderId + "'";
                    db.Database.ExecuteSqlCommand(com);
                    db.Dispose();
                    db = new AppContext();
                    ShowInfo();
                }
                
                
            }
            
        }

        private void Tabitem2_Clicked(object sender, MouseButtonEventArgs e)
        {
            chosen_tab = 2;
            ChangeTabColor();
            ExaminationShortGenerate();
        }
        private void ExaminationShortGenerate()
        {
            lateUsed = false;
            rangeUsed = false;
            EmployeeDataGrid1.AutoGenerateColumns = false;
            var employeeList = from b in db.Employees.ToList()
                               orderby b.OrderId
                               select new Employee
                               {
                                   OrderId = b.OrderId,
                                   Name = b.Name,
                                   Position = b.Position,
                                   Department = b.Department,
                                   ExaminationDateFact = b.ExaminationDateFact,
                                   ExaminationDatePlan = b.ExaminationDatePlan,
                                   ExaminationComplexDateFact = b.ExaminationComplexDateFact,
                                   ExaminationComplexDatePlan = b.ExaminationComplexDatePlan

                               };
            dt = ConvertListToDataTable(employeeList.ToList(), 8);

            EmployeeDataGrid1.ItemsSource = employeeList.ToList();
            

        }
        private void Tabitem4_Clicked(object sender, MouseButtonEventArgs e)
        {

            Tab4Chosen();
        }
        public void Tab4Chosen()
        {
            lateUsed = false;
            rangeUsed = false;
            chosen_tab = 4;
            ChangeTabColor();
            EmployeeDataGrid3.AutoGenerateColumns = false;
            var employeeList = from b in db.Employees.ToList()
                               orderby b.OrderId
                               select new Employee
                               {
                                   OrderId = b.OrderId,
                                   Name = b.Name,
                                   Position = b.Position,
                                   Department = b.Department,
                                   AttestationDateFact = b.AttestationDateFact,
                                   AttestationDatePlan = b.AttestationDatePlan

                               };
            dt = ConvertListToDataTable(employeeList.ToList(), 6);

            EmployeeDataGrid3.ItemsSource = employeeList.ToList();
        }

        private void Tabitem5_Clicked(object sender, MouseButtonEventArgs e)
        {
            Tab5Chosen();
        }
        public void Tab5Chosen()
        {
            lateUsed = false;
            rangeUsed = false;
            chosen_tab = 5;
            ChangeTabColor();
            EmployeeDataGrid4.AutoGenerateColumns = false;
            var employeeList = from b in db.Employees.ToList()
                               orderby b.OrderId
                               select new Employee
                               {
                                   OrderId = b.OrderId,
                                   Name = b.Name,
                                   Position = b.Position,
                                   Department = b.Department,
                                   PbminimumPassDateFact = b.PbminimumPassDateFact,
                                   PbminimumPassDatePlan = b.PbminimumPassDatePlan
                               };
            dt = ConvertListToDataTable(employeeList.ToList(), 6);

            EmployeeDataGrid4.ItemsSource = employeeList.ToList();
        }
        private void Tabitem6_Clicked(object sender, MouseButtonEventArgs e)
        {
            Tab6Chosen();
        }
        public void Tab6Chosen()
        {
            lateUsed = false;
            rangeUsed = false;
            chosen_tab = 6;
            ChangeTabColor();
            EmployeeDataGrid5.AutoGenerateColumns = false;
            var employeeList = from b in db.Employees.ToList()
                               orderby b.OrderId
                               select new Employee
                               {
                                   OrderId = b.OrderId,
                                   Name = b.Name,
                                   Position = b.Position,
                                   Department = b.Department,
                                   TabNumber = b.TabNumber
                               };
            dt = ConvertListToDataTable(employeeList.ToList(), 5);

            EmployeeDataGrid5.ItemsSource = employeeList.ToList();
        }
        private void Tabitem7_Clicked(object sender, MouseButtonEventArgs e)
        {
            Tab7Chosen();
        }
        public void Tab7Chosen()
        {
            lateUsed = false;
            rangeUsed = false;
            chosen_tab = 7;
            ChangeTabColor();
            EmployeeDataGrid6.AutoGenerateColumns = false;
            var employeeList = from b in db.Employees.ToList()
                               orderby b.OrderId
                               select new Employee
                               {
                                   OrderId = b.OrderId,
                                   Name = b.Name,
                                   Position = b.Position,
                                   Department = b.Department,
                                   MedicalCheckDateFact = b.MedicalCheckDateFact,
                                   MedicalCheckDatePlan = b.MedicalCheckDatePlan
                               };
            dt = ConvertListToDataTable(employeeList.ToList(), 6);

            EmployeeDataGrid6.ItemsSource = employeeList.ToList();
        }
        private void Tabitem8_Clicked(object sender, MouseButtonEventArgs e)
        {
            Tab8Chosen();
        }
        public void Tab8Chosen()
        {
            lateUsed = false;
            rangeUsed = false;
            chosen_tab = 8;
            ChangeTabColor();
            EmployeeDataGrid7.AutoGenerateColumns = false;
            var employeeList = from b in db.Employees.ToList()
                               orderby b.OrderId
                               select new Employee
                               {
                                   OrderId = b.OrderId,
                                   Name = b.Name,
                                   Position = b.Position,
                                   Department = b.Department,
                                   InternshipDate = b.InternshipDate,
                                   InternshipDetails = b.InternshipDetails
                               };
            dt = ConvertListToDataTable(employeeList.ToList(), 6);

            EmployeeDataGrid7.ItemsSource = employeeList.ToList();
        }

        private void Tabitem9_Clicked(object sender, MouseButtonEventArgs e)
        {
            Tab9Chosen();
        }
        public void Tab9Chosen()
        {
            lateUsed = false;
            rangeUsed = false;
            chosen_tab = 9;
            ChangeTabColor();
            EmployeeDataGrid8.AutoGenerateColumns = false;
            var employeeList = from b in db.Employees.ToList()
                               orderby b.OrderId
                               select new Employee
                               {
                                   OrderId = b.OrderId,
                                   Name = b.Name,
                                   Position = b.Position,
                                   Department = b.Department,
                                   DublicationDate = b.DublicationDate,
                                   DublicationDetails = b.DublicationDetails
                               };
            dt = ConvertListToDataTable(employeeList.ToList(), 6);

            EmployeeDataGrid8.ItemsSource = employeeList.ToList();
        }
        private void Tabitem10_Clicked(object sender, MouseButtonEventArgs e)
        {
            Tab10Chosen();
        }
        public void Tab10Chosen()
        {
            lateUsed = false;
            rangeUsed = false;
            chosen_tab = 10;
            ChangeTabColor();
            EmployeeDataGrid9.AutoGenerateColumns = false;
            var employeeList = from b in db.Employees.ToList()
                               orderby b.OrderId
                               select new Employee
                               {
                                   OrderId = b.OrderId,
                                   Name = b.Name,
                                   Position = b.Position,
                                   Department = b.Department,
                                   IndependentDate = b.IndependentDate,
                                   IndependentDetails = b.IndependentDetails
                               };
            dt = ConvertListToDataTable(employeeList.ToList(), 6);

            EmployeeDataGrid9.ItemsSource = employeeList.ToList();
        }
        private void Tabitem11_Clicked(object sender, MouseButtonEventArgs e)
        {
            Tab11Chosen();
        }
        public void Tab11Chosen()
        {
            lateUsed = false;
            rangeUsed = false;
            chosen_tab = 11;
            ChangeTabColor();
            EmployeeDataGrid10.AutoGenerateColumns = false;
            var employeeList = from b in db.Employees.ToList()
                               orderby b.OrderId
                               select new Employee
                               {
                                   OrderId = b.OrderId,
                                   Name = b.Name,
                                   Position = b.Position,
                                   Department = b.Department,
                                   EntryDate = b.EntryDate
                               };
            dt = ConvertListToDataTable(employeeList.ToList(), 5);

            EmployeeDataGrid10.ItemsSource = employeeList.ToList();
        }
        private void FormLate_Click(object sender, RoutedEventArgs e)
        {
            lateUsed = true;
            try
            {
                switch (chosen_tab)
                {
                    case 2:
                        var employeeList = from b in db.Employees.ToList()
                                           where b.ExaminationDatePlan != "не требуется"
                                           where b.ExaminationDatePlan == null || b.ExaminationDatePlan == "" || DateTime.Compare(DateTime.Parse(b.ExaminationDatePlan), DateTime.Now) < 0
                                           select new Employee
                                           {
                                               OrderId = b.OrderId,
                                               Name = b.Name,
                                               Position = b.Position,
                                               Department = b.Department,
                                               ExaminationDateFact = b.ExaminationDateFact,
                                               ExaminationDatePlan = b.ExaminationDatePlan,
                                               ExaminationComplexDateFact = b.ExaminationComplexDateFact,
                                               ExaminationComplexDatePlan = b.ExaminationComplexDatePlan
                                           };
                        dt = ConvertListToDataTable(employeeList.ToList(), 8);

                        EmployeeDataGrid1.ItemsSource = employeeList.ToList();
                        break;
                    case 4:
                        var employeeList1 = from b in db.Employees.ToList()
                                            where b.AttestationDatePlan != "не требуется"
                                            where b.AttestationDatePlan == null || b.AttestationDatePlan == "" || DateTime.Compare(DateTime.Parse(b.AttestationDatePlan), DateTime.Now) < 0
                                            select new Employee
                                            {
                                                OrderId = b.OrderId,
                                                Name = b.Name,
                                                Position = b.Position,
                                                Department = b.Department,
                                                AttestationDateFact = b.AttestationDateFact,
                                                AttestationDatePlan = b.AttestationDatePlan
                                            };
                        dt = ConvertListToDataTable(employeeList1.ToList(), 6);

                        EmployeeDataGrid3.ItemsSource = employeeList1.ToList();
                        break;
                    case 5:
                        var employeeList2 = from b in db.Employees.ToList()
                                            where b.PbminimumPassDatePlan != "не требуется"
                                            where b.PbminimumPassDatePlan == null || b.PbminimumPassDatePlan == "" || DateTime.Compare(DateTime.Parse(b.PbminimumPassDatePlan), DateTime.Now) < 0
                                            select new Employee
                                            {
                                                OrderId = b.OrderId,
                                                Name = b.Name,
                                                Position = b.Position,
                                                Department = b.Department,
                                                PbminimumPassDateFact = b.PbminimumPassDateFact,
                                                PbminimumPassDatePlan = b.PbminimumPassDatePlan

                                            };
                        dt = ConvertListToDataTable(employeeList2.ToList(), 6);

                        EmployeeDataGrid4.ItemsSource = employeeList2.ToList();
                        break;
                    default:
                        break;
                }
            }
            catch
            {
                MessageBox.Show("Не получилось обработать все данные, проверьте формат дат", "Ошибка операции", MessageBoxButton.OK, MessageBoxImage.Error);

            }
            
            
        }
        private void FormRange_Click(object sender, RoutedEventArgs e) //вывод персонала для проверки знаний в заданном диапазоне
        {
            rangeUsed = true;
            string from_a = "";
            string to = "";
            try
            {
                switch (chosen_tab)
                {
                    case 2:
                        from_a = DateTime.Parse(from_date.SelectedDate.ToString()).ToString("dd.MM.yyyy");
                        to = DateTime.Parse(to_date.SelectedDate.ToString()).ToString("dd.MM.yyyy");
                        var list = from b in db.Employees.ToList()
                                   where b.ExaminationDatePlan != null && b.ExaminationDatePlan.ToLower() != "не требуется" && b.ExaminationDatePlan != ""
                                   where (DateTime.Compare(DateTime.Parse(b.ExaminationDatePlan), DateTime.Parse(from_a)) > 0) && (DateTime.Compare(DateTime.Parse(b.ExaminationDatePlan), DateTime.Parse(to)) < 0)
                                   select b;
                        dt = ConvertListToDataTable(list.ToList(), 8);
                        EmployeeDataGrid1.ItemsSource = list.ToList();
                        break;
                    
                    case 4:
                        from_a = DateTime.Parse(from_date.SelectedDate.ToString()).ToString("dd.MM.yyyy");
                        to = DateTime.Parse(to_date.SelectedDate.ToString()).ToString("dd.MM.yyyy");
                        var list2 = from b in db.Employees.ToList()
                                    where b.AttestationDatePlan != null && b.AttestationDatePlan.ToLower() != "не требуется" && b.AttestationDatePlan != ""
                                    where (DateTime.Compare(DateTime.Parse(b.AttestationDatePlan), DateTime.Parse(from_a)) > 0) && (DateTime.Compare(DateTime.Parse(b.AttestationDatePlan), DateTime.Parse(to)) < 0)
                                    select b;
                        dt = ConvertListToDataTable(list2.ToList(), 6);
                        EmployeeDataGrid3.ItemsSource = list2.ToList();
                        break;
                    case 5:
                        from_a = DateTime.Parse(from_date.SelectedDate.ToString()).ToString("dd.MM.yyyy");
                        to = DateTime.Parse(to_date.SelectedDate.ToString()).ToString("dd.MM.yyyy");
                        var list3 = from b in db.Employees.ToList()
                                    where b.PbminimumPassDatePlan != null && b.PbminimumPassDatePlan.ToLower() != "не требуется" && b.PbminimumPassDatePlan != ""
                                    where (DateTime.Compare(DateTime.Parse(b.PbminimumPassDatePlan), DateTime.Parse(from_a)) > 0) && (DateTime.Compare(DateTime.Parse(b.PbminimumPassDatePlan), DateTime.Parse(to)) < 0)
                                    select b;
                        dt = ConvertListToDataTable(list3.ToList(), 6);
                        EmployeeDataGrid4.ItemsSource = list3.ToList();
                        break;
                    default:
                        break;
                }
            }
            catch (Exception e3)
            {
                MessageBox.Show("Выберите даты!", "Ошибка операции", MessageBoxButton.OK, MessageBoxImage.Error);
                MessageBox.Show(e3.StackTrace);
            }

        }
        private void SearchString_TextChanged(object sender, TextChangedEventArgs e)
        {
            string text = searchString.Text;

            switch (chosen_tab)
            {
                case 1:
                    SearchInfo(EmployeeDataGrid, 26);
                    break;
                case 2:
                    SearchInfo(EmployeeDataGrid1, 8);
                    break;
                
                case 4:
                    SearchInfo(EmployeeDataGrid3, 6);
                    break;
                case 5:
                    SearchInfo(EmployeeDataGrid4, 6);
                    break;
                case 6:
                    SearchInfo(EmployeeDataGrid5, 5);
                    break;
                case 7:
                    SearchInfo(EmployeeDataGrid6, 6);
                    break;
                case 8:
                    SearchInfo(EmployeeDataGrid7, 6);
                    break;
                case 9:
                    SearchInfo(EmployeeDataGrid8, 6);
                    break;
                case 10:
                    SearchInfo(EmployeeDataGrid9, 6);
                    break;
                case 11:
                    SearchInfo(EmployeeDataGrid10, 5);
                    break;
                default:
                    break;
            }
        }
        private void SearchInfo(DataGrid grid, int columns)
        {
            string text = searchString.Text;
            IEnumerable<Employee> employeeList = null;
            grid.AutoGenerateColumns = false;

            employeeList = from b in db.Employees.ToList()
                           where b.Name != null
                           where b.Name.ToString().ToLower().Contains(text) || b.Name.ToString().Contains(text)
                           select b;
            dt = ConvertListToDataTable(employeeList.ToList(), columns);

            grid.ItemsSource = employeeList.ToList();
        }
        private void AddEmployee_Click(object sender, RoutedEventArgs e)
        {
            switch (chosen_tab)
            {
                case 1:
                    EditCenter(EmployeeDataGrid);
                    DisplayInfoFull();
                    break;
                case 2:
                    EditCenter(EmployeeDataGrid1);
                    ExaminationShortGenerate();
                    break;
               
                case 4:
                    EditCenter(EmployeeDataGrid3);
                    Tab4Chosen();
                    break;
                case 5:
                    EditCenter(EmployeeDataGrid4);
                    Tab5Chosen();
                    break;
                case 6:
                    EditCenter(EmployeeDataGrid5);
                    Tab6Chosen();
                    break;
                case 7:
                    EditCenter(EmployeeDataGrid6);
                    Tab7Chosen();
                    break;
                case 8:
                    EditCenter(EmployeeDataGrid7);
                    Tab8Chosen();
                    break;
                case 9:
                    EditCenter(EmployeeDataGrid8);
                    Tab9Chosen();
                    break;
                case 10:
                    EditCenter(EmployeeDataGrid9);
                    Tab10Chosen();
                    break;
                case 11:
                    EditCenter(EmployeeDataGrid10);
                    Tab11Chosen();
                    break;
                default:
                    break;
            }
        }
        private void EditEmployee_Click(object sender, RoutedEventArgs e)
        {
            string command="";
            Employee employee;
            try
            {
                switch (chosen_tab)
                {
                    case 1:
                        employee = (Employee)EmployeeDataGrid.SelectedItems[0];
                        command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', ExaminationDateFact='" + employee.ExaminationDateFact + "', ExaminationDatePlan='" +
                            employee.ExaminationDatePlan + "', ExaminationComplexDateFact='" + employee.ExaminationComplexDateFact +
                            "', ExaminationComplexDatePlan='" + employee.ExaminationComplexDatePlan + "', AttestationDateFact='" + employee.AttestationDateFact +
                            "', AttestationDatePlan='" + employee.AttestationDatePlan + "', PbminimumPassDateFact='" + employee.PbminimumPassDateFact +
                            "', PbminimumPassDatePlan='" + employee.PbminimumPassDatePlan + "', MedicalCheckDateFact='" + employee.MedicalCheckDateFact +
                            "', MedicalCheckDatePlan='" + employee.MedicalCheckDatePlan + "', TabNumber='" + employee.TabNumber + "', BirthDate='" + employee.BirthDate +
                            "', EntryDate='" + employee.EntryDate + "', RelocationDate='" + employee.RelocationDate + "', PrimaryInstructionDate='" +
                            employee.PrimaryInstructionDate + "', InternshipDate='" + employee.InternshipDate + "', InternshipDetails='" + employee.InternshipDetails + "'," +
                            " DublicationDate='" + employee.DublicationDate + "', DublicationDetails='" + employee.DublicationDetails + "', IndependentDate='" +
                            employee.IndependentDate + "', IndependentDetails='" + employee.IndependentDetails + "', ExtraStatus='" + employee.ExtraStatus + "' WHERE OrderId='" + employee.OrderId + "'";
                        break;
                    case 2:
                        employee = (Employee)EmployeeDataGrid1.SelectedItems[0];
                        command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', ExaminationDateFact='" + employee.ExaminationDateFact + "', ExaminationDatePlan='" +
                            employee.ExaminationDatePlan + "' WHERE OrderId='" + employee.OrderId + "'";
                        break;
                   
                    case 4:
                        employee = (Employee)EmployeeDataGrid3.SelectedItems[0];
                        command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', AttestationDateFact='" + employee.AttestationDateFact +
                            "', AttestationDatePlan='" + employee.AttestationDatePlan + "' WHERE OrderId='" + employee.OrderId + "'";
                        break;
                    case 5:
                        employee = (Employee)EmployeeDataGrid4.SelectedItems[0];
                        command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', PbminimumPassDateFact='" + employee.PbminimumPassDateFact +
                            "', PbminimumPassDatePlan='" + employee.PbminimumPassDatePlan + "' WHERE OrderId='" + employee.OrderId + "'";
                        break;
                    case 6:
                        employee = (Employee)EmployeeDataGrid5.SelectedItems[0];
                        command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', TabNumber=" +employee.TabNumber + "' WHERE OrderId='" + employee.OrderId + "'";
                        break;
                    case 7:
                        employee = (Employee)EmployeeDataGrid6.SelectedItems[0];
                        command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', MedicalCheckDateFact='" + employee.MedicalCheckDateFact +
                            "', MedicalCheckDatePlan='" + employee.MedicalCheckDatePlan + "' WHERE OrderId='" + employee.OrderId + "'";
                        break;
                    case 8:
                        employee = (Employee)EmployeeDataGrid7.SelectedItems[0];
                        command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', InternshipDate='" + employee.InternshipDate +
                            "', InternshipDetails='" + employee.InternshipDetails + "' WHERE OrderId='" + employee.OrderId + "'";
                        break;
                    case 9:
                        employee = (Employee)EmployeeDataGrid8.SelectedItems[0];
                        command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', DublicationDate='" + employee.DublicationDate +
                            "', DublicationDetails='" + employee.DublicationDetails + "' WHERE OrderId='" + employee.OrderId + "'";
                        break;
                    case 10:
                        employee = (Employee)EmployeeDataGrid9.SelectedItems[0];

                        command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', IndependentDate='" + employee.IndependentDate +
                            "', IndependentDetails='" + employee.IndependentDetails + "' WHERE OrderId='" + employee.OrderId + "'";
                        break;
                    case 11:
                        employee = (Employee)EmployeeDataGrid10.SelectedItems[0];
                        command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', EntryDate='" + employee.EntryDate +
                            "' WHERE OrderId='" + employee.OrderId + "'";
                        break;
                }
                db.Database.ExecuteSqlCommand(command);
                db.Dispose();
                db = new AppContext();

            }
            catch (Exception)
            {
                MessageBox.Show("Выберите отредактированную строку для сохранения изменений!", "Ошибка редактирования", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }
        private void DeleteEmployee_Click(object sender, RoutedEventArgs e)
        {
            switch (chosen_tab)
            {
                case 1:
                    DeleteEmployee(EmployeeDataGrid);
                    DisplayInfoFull();
                    break;
                case 2:
                    DeleteEmployee(EmployeeDataGrid1);
                    ExaminationShortGenerate();
                    break;
                
                case 4:
                    DeleteEmployee(EmployeeDataGrid3);
                    Tab4Chosen();
                    break;
                case 5:
                    DeleteEmployee(EmployeeDataGrid4);
                    Tab5Chosen();
                    break;
                case 6:
                    DeleteEmployee(EmployeeDataGrid5);
                    Tab6Chosen();
                    break;
                case 7:
                    DeleteEmployee(EmployeeDataGrid6);
                    Tab7Chosen();
                    break;
                case 8:
                    DeleteEmployee(EmployeeDataGrid7);
                    Tab8Chosen();
                    break;
                case 9:
                    DeleteEmployee(EmployeeDataGrid8);
                    Tab9Chosen();
                    break;
                case 10:
                    DeleteEmployee(EmployeeDataGrid9);
                    Tab10Chosen();
                    break;
                case 11:
                    DeleteEmployee(EmployeeDataGrid10);
                    Tab11Chosen();
                    break;
                default:
                    break;
            }            
        }
        public void DeleteEmployee(DataGrid grid)
        {
            try
            {
                Employee employee = (Employee)grid.SelectedItems[0];
                if (employee.ExaminationDateFact == null)
                    employee.ExaminationDateFact = "";
                if (employee.ExaminationDatePlan == null)
                    employee.ExaminationDatePlan = "";
                if (employee.ExaminationComplexDateFact == null)
                    employee.ExaminationComplexDateFact = "";
                if (employee.ExaminationComplexDatePlan == null)
                    employee.ExaminationComplexDatePlan = "";
                if(employee.AttestationDateFact == null)
                    employee.AttestationDateFact = "";
                if (employee.AttestationDatePlan == null)
                    employee.AttestationDatePlan = "";
                if (employee.PbminimumPassDateFact == null)
                    employee.PbminimumPassDateFact = "";
                if (employee.PbminimumPassDatePlan == null)
                    employee.PbminimumPassDatePlan = "";
                if (employee.MedicalCheckDateFact == null)
                    employee.MedicalCheckDateFact = "";
                if (employee.MedicalCheckDatePlan == null)
                {
                    employee.MedicalCheckDatePlan = "";
                }
                if (employee.BirthDate == null)
                    employee.BirthDate = "";
                if (employee.EntryDate == null)
                {
                    employee.EntryDate = "";
                }
                if (employee.RelocationDate == null)
                    employee.RelocationDate = "";
                if (employee.PrimaryInstructionDate == null)
                    employee.PrimaryInstructionDate = "";
                if (employee.InternshipDate == null)
                    employee.InternshipDate = "";
                if (employee.InternshipDetails == null)
                    employee.InternshipDetails = "";
                if (employee.DublicationDate == null)
                    employee.DublicationDate = "";
                if (employee.DublicationDetails == null)
                    employee.DublicationDetails = "";
                if (employee.IndependentDate == null)
                    employee.IndependentDate = "";
                if (employee.IndependentDetails == null)
                    employee.IndependentDetails = "";
                if (employee.ExtraStatus == null)
                    employee.ExtraStatus = "";
                if (employee.TabNumber == null)
                    employee.TabNumber = 0;
                employee.CellColor = cellColorList[employee.OrderId-1];
                string command = "DELETE FROM Employees WHERE OrderId='" + employee.OrderId + "'";
                MessageBoxResult result = MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение удаления", MessageBoxButton.YesNo, MessageBoxImage.Question);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        operationsDoneDetails.Push(employee);
                        operationsDone.Push("delete");
                        db.Database.ExecuteSqlCommand(command);
                        db.Dispose();
                        db = new AppContext();
                        command = "UPDATE Employees SET OrderId=OrderId-1 WHERE OrderId>" + employee.OrderId;
                        cellColorList.RemoveAt(employee.OrderId - 1);
                        db.Database.ExecuteSqlCommand(command);
                        break;
                }
            }
            catch
            {
                MessageBox.Show("Выберите строку для удаления!", "Ошибка удаления", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                switch (chosen_tab)
                {
                    case 1:
                        DeleteEmployee(EmployeeDataGrid);
                        DisplayInfoFull();
                        break;
                    case 2:
                        DeleteEmployee(EmployeeDataGrid1);
                        ExaminationShortGenerate();
                        break;
                    case 4:
                        DeleteEmployee(EmployeeDataGrid3);
                        Tab4Chosen();
                        break;
                    case 5:
                        DeleteEmployee(EmployeeDataGrid4);
                        Tab5Chosen();
                        break;
                    case 6:
                        DeleteEmployee(EmployeeDataGrid5);
                        Tab6Chosen();
                        break;
                    case 7:
                        DeleteEmployee(EmployeeDataGrid6);
                        Tab7Chosen();
                        break;
                    case 8:
                        DeleteEmployee(EmployeeDataGrid7);
                        Tab8Chosen();
                        break;
                    case 9:
                        DeleteEmployee(EmployeeDataGrid8);
                        Tab9Chosen();
                        break;
                    case 10:
                        DeleteEmployee(EmployeeDataGrid9);
                        Tab10Chosen();
                        break;
                    case 11:
                        DeleteEmployee(EmployeeDataGrid10);
                        Tab11Chosen();
                        break;
                    default:
                        break;
                }

            }
            else if (e.Key==Key.V && Keyboard.Modifiers == ModifierKeys.Control)
            {
                Employee employee = new Employee();
                string data = Clipboard.GetData(DataFormats.Text).ToString();
                string[] cells = data.Split('\t');
                switch (chosen_tab)
                {
                    case 1:
                        if (cells.Length==26)
                        {
                            employee = (Employee)EmployeeDataGrid.SelectedItems[0];
                            employee.Name = cells[1];
                            employee.Position = cells[2];
                            employee.Department = cells[3];
                            employee.ExaminationDateFact = cells[4];
                            employee.ExaminationDatePlan = cells[5];
                            employee.ExaminationComplexDateFact = cells[6];
                            employee.ExaminationComplexDatePlan = cells[7];
                            employee.AttestationDateFact = cells[8];
                            employee.AttestationDatePlan = cells[9];
                            employee.PbminimumPassDateFact = cells[10];
                            employee.PbminimumPassDatePlan = cells[11];
                            employee.MedicalCheckDateFact = cells[12];
                            employee.MedicalCheckDatePlan = cells[13];
                            employee.TabNumber = Convert.ToInt32(cells[14]);
                            employee.BirthDate = cells[15];
                            employee.EntryDate = cells[16];
                            employee.RelocationDate = cells[17];
                            employee.PrimaryInstructionDate = cells[18];
                            employee.InternshipDate = cells[19];
                            employee.InternshipDetails = cells[20];
                            employee.DublicationDate = cells[21];
                            employee.DublicationDetails = cells[22];
                            employee.IndependentDate = cells[23];
                            employee.IndependentDetails = cells[24];
                            employee.ExtraStatus = cells[25];
                            UpdateEmployeeGrid(employee.OrderId - 1);
                            DisplayInfoFull();
                        }
                        
                        
                        break;
                    case 2:
                        if (cells.Length == 8)
                        {
                            employee = (Employee)EmployeeDataGrid1.SelectedItems[0];
                            employee.Name = cells[1];
                            employee.Position = cells[2];
                            employee.Department = cells[3];
                            employee.ExaminationDateFact = cells[4];
                            employee.ExaminationDatePlan = cells[5];
                            employee.ExaminationComplexDateFact = cells[6];
                            employee.ExaminationComplexDatePlan = cells[7];
                            UpdateEmployeeGrid1(employee.OrderId - 1);
                            ExaminationShortGenerate();
                        }
                        
                        break;
                    case 4:
                        if (cells.Length==6)
                        {
                            employee = (Employee)EmployeeDataGrid3.SelectedItems[0];
                            employee.Name = cells[1];
                            employee.Position = cells[2];
                            employee.Department = cells[3];
                            employee.AttestationDateFact = cells[4];
                            employee.AttestationDatePlan = cells[5];
                            UpdateEmployeeGrid3(employee.OrderId - 1);
                            Tab4Chosen();
                        }
                        
                        break;
                    case 5:
                        if (cells.Length==6)
                        {
                            employee = (Employee)EmployeeDataGrid4.SelectedItems[0];
                            employee.Name = cells[1];
                            employee.Position = cells[2];
                            employee.Department = cells[3];
                            employee.PbminimumPassDateFact = cells[4];
                            employee.PbminimumPassDatePlan = cells[5];
                            UpdateEmployeeGrid4(employee.OrderId - 1);
                            Tab5Chosen();
                        }
                        
                        break;
                    case 6:
                        if (cells.Length==5)
                        {
                            employee = (Employee)EmployeeDataGrid5.SelectedItems[0];
                            employee.Name = cells[1];
                            employee.Position = cells[2];
                            employee.Department = cells[3];
                            employee.TabNumber = Convert.ToInt32(cells[4]);
                            UpdateEmployeeGrid5(employee.OrderId - 1);
                            Tab6Chosen();
                        }
                        

                        break;
                    case 7:
                        if (cells.Length==6)
                        {
                            employee = (Employee)EmployeeDataGrid6.SelectedItems[0];
                            employee.Name = cells[1];
                            employee.Position = cells[2];
                            employee.Department = cells[3];
                            employee.MedicalCheckDateFact = cells[4];
                            employee.MedicalCheckDatePlan = cells[5];
                            UpdateEmployeeGrid6(employee.OrderId - 1);
                            Tab7Chosen();
                        }
                        
                        break;
                    case 8:
                        if (cells.Length==6)
                        {
                            employee = (Employee)EmployeeDataGrid7.SelectedItems[0];
                            employee.Name = cells[1];
                            employee.Position = cells[2];
                            employee.Department = cells[3];
                            employee.InternshipDate = cells[4];
                            employee.InternshipDetails = cells[5];
                            UpdateEmployeeGrid7(employee.OrderId - 1);
                            Tab8Chosen();
                        }
                        
                        break;
                    case 9:
                        if (cells.Length==6)
                        {
                            employee = (Employee)EmployeeDataGrid8.SelectedItems[0];
                            employee.Name = cells[1];
                            employee.Position = cells[2];
                            employee.Department = cells[3];
                            employee.DublicationDate = cells[4];
                            employee.DublicationDetails = cells[5];
                            UpdateEmployeeGrid8(employee.OrderId - 1);
                            Tab9Chosen();
                        }
                        
                        break;
                    case 10:
                        if (cells.Length==6)
                        {
                            employee = (Employee)EmployeeDataGrid9.SelectedItems[0];
                            employee.Name = cells[1];
                            employee.Position = cells[2];
                            employee.Department = cells[3];
                            employee.IndependentDate = cells[4];
                            employee.IndependentDetails = cells[5];
                            UpdateEmployeeGrid9(employee.OrderId - 1);
                            Tab10Chosen();
                        }
                        
                        break;
                    case 11:
                        if (cells.Length==5)
                        {
                            employee = (Employee)EmployeeDataGrid10.SelectedItems[0];
                            employee.Name = cells[1];
                            employee.Position = cells[2];
                            employee.Department = cells[3];
                            employee.EntryDate = cells[4];
                            UpdateEmployeeGrid10(employee.OrderId - 1);
                            Tab11Chosen();
                        }
                        
                        break;


                }
            }
        }
        private void SaveAsExcel_Click(object sender, RoutedEventArgs e)
        {
            
            APP = new Microsoft.Office.Interop.Excel.Application();
            OpenFileDialog folderBrowser = new OpenFileDialog();
            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            folderBrowser.FileName = "";
            string path = "";
            try
            {
                Workbook book = SavePrintAction();

                if (folderBrowser.ShowDialog() == true)
                {
                    path = folderBrowser.FileName;
                    if (!path.EndsWith(".xlsx"))
                        path += ".xlsx";

                }
                book.SaveToFile(path);
                System.Diagnostics.Process.Start(path);
                MessageBox.Show("Таблица сохранена в excel файл!", "Результат сохранения", MessageBoxButton.OK, MessageBoxImage.Information);
                
                
            }
            catch
            {
                MessageBox.Show("Ошибка операции", "Статус операции", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }
        public Workbook SavePrintAction()
        {
            Workbook book = new Workbook();
            Worksheet sheet = book.Worksheets[0];
            DataTable temp = dt;
            int j = 0;
            sheet.InsertDataTable(temp, true, 1, 1);
            sheet.Range["A1:Z400"].Style.WrapText = true;
            sheet.Range["A1:Z1"].Style.HorizontalAlignment = Spire.Xls.HorizontalAlignType.Center;
            sheet.Range["A1:Z400"].Style.VerticalAlignment = Spire.Xls.VerticalAlignType.Top;
            sheet.Range["A2:Z400"].Style.HorizontalAlignment = Spire.Xls.HorizontalAlignType.Left;
            sheet.Range["A1:Z400"].Style.Font.FontName = "Times New Roman";
            sheet.Range["A1:Z400"].Style.Font.Size = 12;
            sheet.Range["D2:D400"].Style.Font.IsBold = true;
            sheet.Range["A2:A400"].IgnoreErrorOptions = IgnoreErrorType.All;
            sheet.Range["O2:O400"].IgnoreErrorOptions = IgnoreErrorType.All;
            sheet.Range["P2:P400"].IgnoreErrorOptions = IgnoreErrorType.All;
            switch (chosen_tab)
            {
                case 1:
                    for (int i = 0; i < cellColorList.Count; i++)
                    {
                        Color color = (Color)ColorConverter.ConvertFromString(cellColorList[i]);
                        sheet.Range["A1:Z1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                        sheet.Range[$"A1:Z{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);

                        sheet.Range[$"A{i + 2}:Z{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);

                    }
                    break;
                case 2:
                    if (!lateUsed&&!rangeUsed)
                    {
                        sheet.Range["A1:H1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                        sheet.Range[$"A1:H{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);
                        for (int i = 0; i < cellColorList.Count; i++)
                        {
                            Color color = (Color)ColorConverter.ConvertFromString(cellColorList[i]);
                            sheet.Range[$"A{i + 2}:H{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);

                        }
                    }
                    else
                    {
                        sheet.Range["A1:H1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                        sheet.Range[$"A1:H{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);
                        for (int i = 0; i<dt.Rows.Count; i++)
                        {
                            j = Convert.ToInt32(dt.Rows[i][0].ToString());
                            Color color = (Color)ColorConverter.ConvertFromString(cellColorList[j-1]);

                            sheet.Range[$"A{i + 2}:H{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);
                        }
                    }
                    
                    break;
                case 4:
                    if (!lateUsed && !rangeUsed)
                    {
                        for (int i = 0; i < cellColorList.Count; i++)
                        {
                            Color color = (Color)ColorConverter.ConvertFromString(cellColorList[i]);
                            sheet.Range["A1:F1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                            sheet.Range[$"A1:F{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);

                            sheet.Range[$"A{i + 2}:F{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);

                        }
                    }
                    else
                    {
                        sheet.Range["A1:F1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                        sheet.Range[$"A1:F{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            j = Convert.ToInt32(dt.Rows[i][0].ToString());
                            Color color = (Color)ColorConverter.ConvertFromString(cellColorList[j - 1]);

                            sheet.Range[$"A{i + 2}:F{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);
                        }
                    }
                    break;
                case 5:
                    if (!lateUsed && !rangeUsed)
                    {

                        sheet.Range["A1:F1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                        sheet.Range[$"A1:F{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);
                        for (int i = 0; i < cellColorList.Count; i++)
                        {
                            Color color = (Color)ColorConverter.ConvertFromString(cellColorList[i]);

                            sheet.Range[$"A{i + 2}:F{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);

                        }
                    }
                    else
                    {
                        sheet.Range["A1:F1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                        sheet.Range[$"A1:F{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            j = Convert.ToInt32(dt.Rows[i][0].ToString());
                            Color color = (Color)ColorConverter.ConvertFromString(cellColorList[j - 1]);

                            sheet.Range[$"A{i + 2}:F{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);
                        }
                    }
                    break;
                case 6:
                    sheet.Range["A1:E1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                    sheet.Range[$"A1:E{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);

                    for (int i = 0; i < cellColorList.Count; i++)
                    {
                        Color color = (Color)ColorConverter.ConvertFromString(cellColorList[i]);
                        
                        sheet.Range[$"A{i + 2}:E{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);

                    }
                    break;
                case 7:
                    sheet.Range["A1:F1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                    sheet.Range[$"A1:F{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);

                    for (int i = 0; i < cellColorList.Count; i++)
                    {
                        Color color = (Color)ColorConverter.ConvertFromString(cellColorList[i]);
                        
                        sheet.Range[$"A{i + 2}:F{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);

                    }
                    break;
                case 8:
                    sheet.Range["A1:F1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                    sheet.Range[$"A1:F{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);

                    for (int i = 0; i < cellColorList.Count; i++)
                    {
                        Color color = (Color)ColorConverter.ConvertFromString(cellColorList[i]);
                        
                        sheet.Range[$"A{i + 2}:F{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);

                    }
                    break;
                case 9:
                    sheet.Range["A1:F1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                    sheet.Range[$"A1:F{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);

                    for (int i = 0; i < cellColorList.Count; i++)
                    {
                        Color color = (Color)ColorConverter.ConvertFromString(cellColorList[i]);
                        
                        sheet.Range[$"A{i + 2}:F{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);

                    }
                    break;
                case 10:
                    sheet.Range["A1:F1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                    sheet.Range[$"A1:F{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);

                    for (int i = 0; i < cellColorList.Count; i++)
                    {
                        Color color = (Color)ColorConverter.ConvertFromString(cellColorList[i]);
                        
                        sheet.Range[$"A{i + 2}:F{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);

                    }
                    break;
                case 11:
                    sheet.Range["A1:E1"].Style.Color = System.Drawing.Color.FromArgb(208, 206, 206);
                    sheet.Range[$"A1:E{dt.Rows.Count+1}"].BorderInside(LineStyleType.Thin, System.Drawing.Color.Black);
                    

                    for (int i = 0; i < cellColorList.Count; i++)
                    {
                        Color color = (Color)ColorConverter.ConvertFromString(cellColorList[i]);
                        
                        sheet.Range[$"A{i + 2}:E{i + 2}"].Style.Color = System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);

                    }
                    break;
            }
            CreateHeader(sheet);
            return book;

        }
        private void CreateHeader(Worksheet sheet)
        {
            int i = 0;
            List<string> columns = new List<string> { "A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1", "J1", "K1", "L1", "M1", "N1", "O1", "P1", "Q1", "R1", "S1", "T1", "U1", "V1", "W1", "X1", "Y1", "Z1" };
            int[] columnLength;
            switch (chosen_tab)
            {
                case 1:
                    columnLength = new int[] { 6, 20, 21, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 8, 8, 13, 13, 10, 13, 11, 13, 13, 13, 13, 13, 11 };

                    foreach (object ob in EmployeeDataGrid.Columns.Select(cs => cs.Header).ToList())
                    {
                        //WS.Cells[1, ind] = ob.ToString();
                        //WS.Cells.WrapText = true;
                        //WS.Cells[1, ind].ColumnWidth = columnLength[ind - 1];
                        //WS.Cells[1, ind].Interior.Color = XlRgbColor.rgbBeige;
                        //ind++;
                        sheet[columns[i]].Text = ob.ToString();
                        sheet[columns[i]].ColumnWidth = columnLength[i];
                        i++;
                    }
                    break;
                case 2:
                    columnLength = new int[] { 6, 20, 21, 13, 13, 13, 13, 13 };
                    foreach (object ob in EmployeeDataGrid1.Columns.Select(cs => cs.Header).ToList())
                    {
                        //WS.Cells[1, ind] = ob.ToString();
                        //WS.Cells.WrapText = true;
                        //WS.Cells[1, ind].ColumnWidth = columnLength[ind - 1];
                        //WS.Cells[1, ind].Interior.Color = XlRgbColor.rgbBeige;
                        //ind++;
                        sheet[columns[i]].Text = ob.ToString();
                        sheet[columns[i]].ColumnWidth = columnLength[i];

                        i++;
                    }
                    break;
                //case 3:
                //    columnLength = new int[] { 6, 20, 21, 13, 13, 13 };
                //    foreach (object ob in EmployeeDataGrid2.Columns.Select(cs => cs.Header).ToList())
                //    {
                //        WS.Cells[1, ind] = ob.ToString();
                //        WS.Cells.WrapText = true;
                //        WS.Cells[1, ind].ColumnWidth = columnLength[ind - 1];
                //        WS.Cells[1, ind].Interior.Color = XlRgbColor.rgbBeige;
                //        ind++;
                //    }
                //    break;
                case 4:
                    columnLength = new int[] { 6, 20, 21, 13, 13, 13 };
                    foreach (object ob in EmployeeDataGrid3.Columns.Select(cs => cs.Header).ToList())
                    {
                        //WS.Cells[1, ind] = ob.ToString();
                        //WS.Cells.WrapText = true;
                        //WS.Cells[1, ind].ColumnWidth = columnLength[ind - 1];
                        //WS.Cells[1, ind].Interior.Color = XlRgbColor.rgbBeige;
                        //ind++;
                        sheet[columns[i]].Text = ob.ToString();
                        sheet[columns[i]].ColumnWidth = columnLength[i];

                        i++;
                    }
                    break;
                case 5:
                    columnLength = new int[] { 6, 20, 21, 13, 13, 13 };
                    foreach (object ob in EmployeeDataGrid4.Columns.Select(cs => cs.Header).ToList())
                    {
                        //WS.Cells[1, ind] = ob.ToString();
                        //WS.Cells.WrapText = true;
                        //WS.Cells[1, ind].ColumnWidth = columnLength[ind - 1];
                        //WS.Cells[1, ind].Interior.Color = XlRgbColor.rgbBeige;
                        //ind++;
                        sheet[columns[i]].Text = ob.ToString();
                        sheet[columns[i]].ColumnWidth = columnLength[i];

                        i++;
                    }
                    break;
                case 6:
                    columnLength = new int[] { 6, 20, 21, 13, 13 };
                    foreach (object ob in EmployeeDataGrid5.Columns.Select(cs => cs.Header).ToList())
                    {
                        //WS.Cells[1, ind] = ob.ToString();
                        //WS.Cells.WrapText = true;
                        //WS.Cells[1, ind].ColumnWidth = columnLength[ind - 1];
                        //WS.Cells[1, ind].Interior.Color = XlRgbColor.rgbBeige;
                        //ind++;
                        sheet[columns[i]].Text = ob.ToString();
                        sheet[columns[i]].ColumnWidth = columnLength[i];

                        i++;
                    }
                    break;
                case 7:
                    columnLength = new int[] { 6, 20, 21, 13, 13, 13 };
                    foreach (object ob in EmployeeDataGrid6.Columns.Select(cs => cs.Header).ToList())
                    {
                        //WS.Cells[1, ind] = ob.ToString();
                        //WS.Cells.WrapText = true;
                        //WS.Cells[1, ind].ColumnWidth = columnLength[ind - 1];
                        //WS.Cells[1, ind].Interior.Color = XlRgbColor.rgbBeige;
                        //ind++;
                        sheet[columns[i]].Text = ob.ToString();
                        sheet[columns[i]].ColumnWidth = columnLength[i];

                        i++;
                    }
                    break;
                case 8:
                    columnLength = new int[] { 6, 20, 21, 13, 13, 13 };
                    foreach (object ob in EmployeeDataGrid7.Columns.Select(cs => cs.Header).ToList())
                    {
                        //WS.Cells[1, ind] = ob.ToString();
                        //WS.Cells.WrapText = true;
                        //WS.Cells[1, ind].ColumnWidth = columnLength[ind - 1];
                        //WS.Cells[1, ind].Interior.Color = XlRgbColor.rgbBeige;
                        //ind++;
                        sheet[columns[i]].Text = ob.ToString();
                        sheet[columns[i]].ColumnWidth = columnLength[i];

                        i++;
                    }
                    break;
                case 9:
                    columnLength = new int[] { 6, 20, 21, 13, 13, 13 };
                    foreach (object ob in EmployeeDataGrid8.Columns.Select(cs => cs.Header).ToList())
                    {
                        //WS.Cells[1, ind] = ob.ToString();
                        //WS.Cells.WrapText = true;
                        //WS.Cells[1, ind].ColumnWidth = columnLength[ind - 1];
                        //WS.Cells[1, ind].Interior.Color = XlRgbColor.rgbBeige;
                        //ind++;
                        sheet[columns[i]].Text = ob.ToString();
                        sheet[columns[i]].ColumnWidth = columnLength[i];

                        i++;
                    }
                    break;
                case 10:
                    columnLength = new int[] { 6, 20, 21, 13, 13, 13 };
                    foreach (object ob in EmployeeDataGrid9.Columns.Select(cs => cs.Header).ToList())
                    {
                        //WS.Cells[1, ind] = ob.ToString();
                        //WS.Cells.WrapText = true;
                        //WS.Cells[1, ind].ColumnWidth = columnLength[ind - 1];
                        //WS.Cells[1, ind].Interior.Color = XlRgbColor.rgbBeige;
                        //ind++;
                        sheet[columns[i]].Text = ob.ToString();
                        sheet[columns[i]].ColumnWidth = columnLength[i];

                        i++;
                    }
                    break;
                case 11:
                    columnLength = new int[] { 6, 20, 21, 13, 13 };
                    foreach (object ob in EmployeeDataGrid10.Columns.Select(cs => cs.Header).ToList())
                    {
                        //WS.Cells[1, ind] = ob.ToString();
                        //WS.Cells.WrapText = true;
                        //WS.Cells[1, ind].ColumnWidth = columnLength[ind - 1];
                        //WS.Cells[1, ind].Interior.Color = XlRgbColor.rgbBeige;
                        //ind++;
                        sheet[columns[i]].Text = ob.ToString();
                        sheet[columns[i]].ColumnWidth = columnLength[i];

                        i++;
                    }
                    break;
                default:
                    break;
            }

        }
        private void EmployeeDataGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            int b = e.Row.GetIndex();
            UpdateEmployeeGrid(b);
        }
        public void UpdateEmployeeGrid(int b)
        {
            string command = "";
            Employee employee;
            bool canUpdate = false;
            DataGridRow row;

            row = (DataGridRow)EmployeeDataGrid.ItemContainerGenerator.ContainerFromIndex(b);
            employee = (Employee)row.Item;
            if (db.Employees.Find(employee.OrderId) != null)
            {
                command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                    employee.Department + "', ExaminationDateFact='" + employee.ExaminationDateFact + "', ExaminationDatePlan='" +
                    employee.ExaminationDatePlan + "', ExaminationComplexDateFact='" + employee.ExaminationComplexDateFact +
                    "', ExaminationComplexDatePlan='" + employee.ExaminationComplexDatePlan + "', AttestationDateFact='" + employee.AttestationDateFact +
                    "', AttestationDatePlan='" + employee.AttestationDatePlan + "', PbminimumPassDateFact='" + employee.PbminimumPassDateFact +
                    "', PbminimumPassDatePlan='" + employee.PbminimumPassDatePlan + "', MedicalCheckDateFact='" + employee.MedicalCheckDateFact +
                    "', MedicalCheckDatePlan='" + employee.MedicalCheckDatePlan + "', TabNumber='" + employee.TabNumber + "', BirthDate='" + employee.BirthDate +
                    "', EntryDate='" + employee.EntryDate + "', RelocationDate='" + employee.RelocationDate + "', PrimaryInstructionDate='" +
                    employee.PrimaryInstructionDate + "', InternshipDate='" + employee.InternshipDate + "', InternshipDetails='" + employee.InternshipDetails + "'," +
                    " DublicationDate='" + employee.DublicationDate + "', DublicationDetails='" + employee.DublicationDetails + "', IndependentDate='" +
                    employee.IndependentDate + "', IndependentDetails='" + employee.IndependentDetails + "', ExtraStatus='" + employee.ExtraStatus + "' WHERE OrderId='" + employee.OrderId + "'";
                canUpdate = true;
            }
            else
            {
                db.Employees.Add(employee);
                db.SaveChanges();
            }

            if (canUpdate)
            {
                db1 = new AppContext();
                db1 = db;
                db.Database.ExecuteSqlCommand(command);
                //db.Dispose();
                db = new AppContext();

            }
            else
            {
                MessageBox.Show("Данные успешно добавлены!", "Результат добавления", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            DisplayInfoFull();


        }
        private void EmployeeDataGrid1_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            int b = e.Row.GetIndex();
            UpdateEmployeeGrid1(b);
            
        }
        public void UpdateEmployeeGrid1(int b)
        {
            string command = "";
            Employee employee;
            bool canUpdate = false;
            DataGridRow row;
            row = (DataGridRow)EmployeeDataGrid1.ItemContainerGenerator.ContainerFromIndex(b);
            employee = (Employee)row.Item;
            if (db.Employees.Find(employee.OrderId) != null)
            {
                command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                employee.Department + "', ExaminationDateFact='" + employee.ExaminationDateFact + "', ExaminationDatePlan='" +
                employee.ExaminationDatePlan + "', ExaminationComplexDateFact='" + employee.ExaminationComplexDateFact + "', " +
                "ExaminationComplexDatePlan='" + employee.ExaminationComplexDatePlan + "' WHERE OrderId='" + employee.OrderId + "'";
                canUpdate = true;
            }
            else
            {
                db.Employees.Add(employee);
                db.SaveChanges();
            }
            if (canUpdate)
            {
                db.Database.ExecuteSqlCommand(command);
                db.Dispose();
                db = new AppContext();
            }
            else
            {
                MessageBox.Show("Данные успешно добавлены!", "Результат добавления", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            ExaminationShortGenerate();

        }
        private void EmployeeDataGrid3_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            int b = e.Row.GetIndex();
            UpdateEmployeeGrid3(b);
        }
        public void UpdateEmployeeGrid3(int b)
        {
            string command = "";
            Employee employee;
            bool canUpdate = false;
            DataGridRow row;
            row = (DataGridRow)EmployeeDataGrid3.ItemContainerGenerator.ContainerFromIndex(b);
            employee = (Employee)row.Item;
            if (db.Employees.Find(employee.OrderId) != null)
            {
                command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                         employee.Department + "', AttestationDateFact='" + employee.AttestationDateFact +
                         "', AttestationDatePlan='" + employee.AttestationDatePlan + "' WHERE OrderId='" + employee.OrderId + "'";
                canUpdate = true;
            }
            else
            {
                db.Employees.Add(employee);
                db.SaveChanges();
            }
            if (canUpdate)
            {
                db.Database.ExecuteSqlCommand(command);
                db.Dispose();
                db = new AppContext();

            }
            else
            {
                MessageBox.Show("Данные успешно добавлены!", "Результат добавления", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            Tab4Chosen();
        }
        private void EmployeeDataGrid4_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            int b = e.Row.GetIndex();
            UpdateEmployeeGrid4(b);
        }
        public void UpdateEmployeeGrid4(int b)
        {
            string command = "";
            Employee employee;
            bool canUpdate = false;
            DataGridRow row;
            row = (DataGridRow)EmployeeDataGrid4.ItemContainerGenerator.ContainerFromIndex(b);
            employee = (Employee)row.Item;
            if (db.Employees.Find(employee.OrderId) != null)
            {
                command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                         employee.Department + "', PbminimumPassDateFact='" + employee.PbminimumPassDateFact +
                         "', PbminimumPassDatePlan='" + employee.PbminimumPassDatePlan + "' WHERE OrderId='" + employee.OrderId + "'";
                canUpdate = true;
            }
            else
            {
                db.Employees.Add(employee);
                db.SaveChanges();
            }
            if (canUpdate)
            {
                db.Database.ExecuteSqlCommand(command);
                db.Dispose();
                db = new AppContext();

            }
            else
            {
                MessageBox.Show("Данные успешно добавлены!", "Результат добавления", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            Tab5Chosen();
        }
        private void EmployeeDataGrid5_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            int b = e.Row.GetIndex();
            UpdateEmployeeGrid5(b);
        }
        public void UpdateEmployeeGrid5(int b)
        {
            string command = "";
            Employee employee;
            bool canUpdate = false;
            DataGridRow row;
            row = (DataGridRow)EmployeeDataGrid5.ItemContainerGenerator.ContainerFromIndex(b);
            employee = (Employee)row.Item;
            if (db.Employees.Find(employee.OrderId) != null)
            {
                command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', TabNumber='" + employee.TabNumber + "' WHERE OrderId='" + employee.OrderId + "'";
                canUpdate = true;
            }
            else
            {
                db.Employees.Add(employee);
                db.SaveChanges();
            }
            if (canUpdate)
            {
                db.Database.ExecuteSqlCommand(command);
                db.Dispose();
                db = new AppContext();
            }
            Tab6Chosen();

        }
        private void EmployeeDataGrid6_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            int b = e.Row.GetIndex();
            UpdateEmployeeGrid6(b);
        }
        public void UpdateEmployeeGrid6(int b)
        {
            string command = "";
            Employee employee;
            bool canUpdate = false;
            DataGridRow row;
            row = (DataGridRow)EmployeeDataGrid6.ItemContainerGenerator.ContainerFromIndex(b);
            employee = (Employee)row.Item;
            if (db.Employees.Find(employee.OrderId) != null)
            {
                command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                             employee.Department + "', MedicalCheckDateFact='" + employee.MedicalCheckDateFact +
                             "', MedicalCheckDatePlan='" + employee.MedicalCheckDatePlan + "' WHERE OrderId='" + employee.OrderId + "'";
                canUpdate = true;
            }
            else
            {
                db.Employees.Add(employee);
                db.SaveChanges();
            }
            if (canUpdate)
            {
                db.Database.ExecuteSqlCommand(command);
                db.Dispose();
                db = new AppContext();
            }
            Tab7Chosen();
        }
        private void EmployeeDataGrid7_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            int b = e.Row.GetIndex();
            UpdateEmployeeGrid7(b);
        }
        public void UpdateEmployeeGrid7(int b)
        {
            string command = "";
            Employee employee;
            bool canUpdate = false;
            DataGridRow row;
            row = (DataGridRow)EmployeeDataGrid7.ItemContainerGenerator.ContainerFromIndex(b);
            employee = (Employee)row.Item;
            if (db.Employees.Find(employee.OrderId) != null)
            {
                command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', InternshipDate='" + employee.InternshipDate +
                            "', InternshipDetails='" + employee.InternshipDetails + "' WHERE OrderId='" + employee.OrderId + "'";
                canUpdate = true;
            }
            else
            {
                db.Employees.Add(employee);
                db.SaveChanges();
            }
            if (canUpdate)
            {
                db.Database.ExecuteSqlCommand(command);
                db.Dispose();
                db = new AppContext();
            }
            Tab8Chosen();
        }
        private void EmployeeDataGrid8_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            int b = e.Row.GetIndex();
            UpdateEmployeeGrid8(b);
        }
        public void UpdateEmployeeGrid8(int b)
        {
            string command = "";
            Employee employee;
            bool canUpdate = false;
            DataGridRow row;
            row = (DataGridRow)EmployeeDataGrid8.ItemContainerGenerator.ContainerFromIndex(b);
            employee = (Employee)row.Item;
            if (db.Employees.Find(employee.OrderId) != null)
            {
                command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', DublicationDate='" + employee.DublicationDate +
                            "', DublicationDetails='" + employee.DublicationDetails + "' WHERE OrderId='" + employee.OrderId + "'";
                canUpdate = true;
            }
            else
            {
                db.Employees.Add(employee);
                db.SaveChanges();
            }
            if (canUpdate)
            {
                db.Database.ExecuteSqlCommand(command);
                db.Dispose();
                db = new AppContext();
            }
            Tab9Chosen();
        }
        private void EmployeeDataGrid9_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            int b = e.Row.GetIndex();
            UpdateEmployeeGrid9(b);
        }
        public void UpdateEmployeeGrid9(int b)
        {
            string command = "";
            Employee employee;
            bool canUpdate = false;
            DataGridRow row;
            row = (DataGridRow)EmployeeDataGrid9.ItemContainerGenerator.ContainerFromIndex(b);
            employee = (Employee)row.Item;
            if (db.Employees.Find(employee.OrderId) != null)
            {
                command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', IndependentDate='" + employee.IndependentDate +
                            "', IndependentDetails='" + employee.IndependentDetails + "' WHERE OrderId='" + employee.OrderId + "'";
                canUpdate = true;
            }
            else
            {
                db.Employees.Add(employee);
                db.SaveChanges();
            }
            if (canUpdate)
            {
                db.Database.ExecuteSqlCommand(command);
                db.Dispose();
                db = new AppContext();
            }
            Tab10Chosen();
        }
        private void EmployeeDataGrid10_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            int b = e.Row.GetIndex();
            UpdateEmployeeGrid10(b);
        }
        public void UpdateEmployeeGrid10(int b)
        {
            string command = "";
            Employee employee;
            bool canUpdate = false;
            DataGridRow row;
            row = (DataGridRow)EmployeeDataGrid10.ItemContainerGenerator.ContainerFromIndex(b);
            employee = (Employee)row.Item;
            if (db.Employees.Find(employee.OrderId) != null)
            {
                command = "UPDATE Employees SET Name='" + employee.Name + "', Position='" + employee.Position + "', Department='" +
                            employee.Department + "', EntryDate='" + employee.EntryDate +
                            "' WHERE OrderId='" + employee.OrderId + "'";
                canUpdate = true;
            }
            else
            {
                db.Employees.Add(employee);
                db.SaveChanges();
            }
            if (canUpdate)
            {
                db.Database.ExecuteSqlCommand(command);
                db.Dispose();
                db = new AppContext();
            }
            Tab11Chosen();
        }
        public void EditCenter(DataGrid grid)
        {
            try
            {
                Employee employee = (Employee)grid.SelectedItems[0];
                operationsDoneDetails.Push(employee);
                operationsDone.Push("insert");
                string com;

                com = "UPDATE Employees SET OrderId=OrderId+1 WHERE OrderId>=" + employee.OrderId;
                db.Database.ExecuteSqlCommand(com);
                cellColorList.Insert(employee.OrderId-1, "#FFFFFFFF");
                com = "INSERT INTO Employees (OrderId, TabNumber, CellColor) VALUES (" + employee.OrderId + ",0, '#FFFFFFFF')";
                db.Database.ExecuteSqlCommand(com);


                db.Dispose();
                db = new AppContext();


            }
            catch(Exception ex)
            {
                MessageBox.Show("Выберите строку, под которой хотите добавить новую", "Ошибка добавления", MessageBoxButton.OK, MessageBoxImage.Error);
                Console.WriteLine(ex.StackTrace + "\n" + ex.Message);
            }

            
        }
        public DataTable ConvertListToDataTable(List<Employee> list, int columns)
        {
            // New table.
            DataTable table = new DataTable();
            // Get max columns.
            List<string> columnList = new List<string>();
            columnList.Add("OrderId");
            columnList.Add("Name");
            columnList.Add("Position");
            columnList.Add("Department");
            columnList.Add("ExaminationDateFact");
            columnList.Add("ExaminationDatePlan");
            columnList.Add("ExaminationComplexDateFact");
            columnList.Add("ExaminationComplexDatePlan");
            columnList.Add("AttestationDateFact");
            columnList.Add("AttestationDatePlan");
            columnList.Add("PbminimumPassDateFact");
            columnList.Add("PbminimumPassDatePlan");
            columnList.Add("MedicalCheckDateFact");
            columnList.Add("MedicalCheckDatePlan");
            columnList.Add("TabNumber");
            columnList.Add("BirthDate");
            columnList.Add("EntryDate");
            columnList.Add("RelocationDate");
            columnList.Add("PrimaryInstructionDate");
            columnList.Add("InternshipDate");
            columnList.Add("InternshipDetails");
            columnList.Add("DublicationDate");
            columnList.Add("DublicationDetails");
            columnList.Add("IndependentDate");
            columnList.Add("IndependentDetails");
            columnList.Add("ExtraStatus");

            
            switch (chosen_tab)
            {
                case 1:
                    
                    for (int i = 0; i < columns; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];

                    }
                    foreach (var array in list)
                    {
                        table.Rows.Add(array.OrderId, array.Name, array.Position, array.Department, array.ExaminationDateFact, array.ExaminationDatePlan, array.ExaminationComplexDateFact,
                            array.ExaminationComplexDatePlan, array.AttestationDateFact, array.AttestationDatePlan, array.PbminimumPassDateFact, array.PbminimumPassDatePlan,
                            array.MedicalCheckDateFact, array.MedicalCheckDatePlan, array.TabNumber, array.BirthDate, array.EntryDate, array.RelocationDate,
                            array.PrimaryInstructionDate, array.InternshipDate, array.InternshipDetails, array.DublicationDate, array.DublicationDetails, array.IndependentDate,
                            array.IndependentDetails, array.ExtraStatus);
                    }
                    
                    break;
                case 2:
                    for (int i = 0; i < columns; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];

                    }
                    foreach (var array in list)
                    {
                        table.Rows.Add(array.OrderId, array.Name, array.Position, array.Department, array.ExaminationDateFact, array.ExaminationDatePlan, array.ExaminationComplexDateFact, array.ExaminationComplexDatePlan);

                    }
                    break;
                case 3:
                    for (int i = 0; i < 4; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];

                    }
                    for (int i = 6; i < 8; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];
                    }
                    foreach (var array in list)
                    {
                        table.Rows.Add(array.OrderId, array.Name, array.Position, array.Department, array.ExaminationComplexDateFact, array.ExaminationComplexDatePlan);

                    }
                    break;
                case 4:
                    for (int i = 0; i < 4; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];

                    }
                    for (int i = 8; i < 10; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];
                    }
                    foreach (var array in list)
                    {
                        table.Rows.Add(array.OrderId, array.Name, array.Position, array.Department, array.AttestationDateFact, array.AttestationDatePlan);

                    }
                    break;
                case 5:
                    for (int i = 0; i < 4; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];

                    }
                    for (int i = 10; i < 12; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];
                    }
                    foreach (var array in list)
                    {
                        table.Rows.Add(array.OrderId, array.Name, array.Position, array.Department, array.PbminimumPassDateFact, array.PbminimumPassDatePlan);

                    }
                    break;
                case 6:
                    for (int i = 0; i < 4; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];

                    }
                    for (int i = 14; i < 15; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];
                    }
                    foreach (var array in list)
                    {
                        table.Rows.Add(array.OrderId, array.Name, array.Position, array.Department, array.TabNumber);

                    }
                    break;
                case 7:
                    for (int i = 0; i < 4; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];

                    }
                    for (int i = 12; i < 14; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];
                    }
                    foreach (var array in list)
                    {
                        table.Rows.Add(array.OrderId, array.Name, array.Position, array.Department, array.MedicalCheckDateFact, array.MedicalCheckDatePlan);

                    }
                    break;
                case 8:
                    for (int i = 0; i < 4; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];

                    }
                    for (int i = 19; i < 21; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];
                    }
                    foreach (var array in list)
                    {
                        table.Rows.Add(array.OrderId, array.Name, array.Position, array.Department, array.InternshipDate, array.InternshipDetails);

                    }
                    break;
                case 9:
                    for (int i = 0; i < 4; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];

                    }
                    for (int i = 21; i < 23; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];
                    }
                    foreach (var array in list)
                    {
                        table.Rows.Add(array.OrderId, array.Name, array.Position, array.Department, array.DublicationDate, array.DublicationDetails);

                    }
                    break;
                case 10:
                    for (int i = 0; i < 4; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];

                    }
                    for (int i = 23; i < 25; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];
                    }
                    foreach (var array in list)
                    {
                        table.Rows.Add(array.OrderId, array.Name, array.Position, array.Department, array.IndependentDate, array.IndependentDetails);

                    }
                    break;
                case 11:
                    for (int i = 0; i < 4; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];

                    }
                    for (int i = 16; i < 17; i++)
                    {
                        DataColumn column = table.Columns.Add();
                        column.ColumnName = columnList[i];
                    }
                    foreach (var array in list)
                    {
                        table.Rows.Add(array.OrderId, array.Name, array.Position, array.Department, array.EntryDate);

                    }
                    break;
                default:
                    break;
            }
            

            return table;
        }


        private void PrintData_Click(object sender, RoutedEventArgs e) //печать таблицы
        {
            printing = 1;
            APP = new Microsoft.Office.Interop.Excel.Application();
            OpenFileDialog folderBrowser = new OpenFileDialog();

            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            folderBrowser.FileName = "";
            try
            {
                Workbook book = SavePrintAction();

                if (printing == 1)
                {
                    PrintDialog dialogPrint = new PrintDialog();
                    if (dialogPrint.ShowDialog() == true)
                    {
                        book.PrintDocument.Print();
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка печати", MessageBoxButton.OK, MessageBoxImage.Error);
            }
                //try
                //{
                //    if (printDialog.ShowDialog() == true)
                //    {
                //        switch (chosen_tab)
                //        {
                //            case 1:
                //                printDialog.PrintVisual(EmployeeDataGrid, "Печать таблицы");
                //                break;
                //            case 2:
                //                printDialog.PrintVisual(EmployeeDataGrid1, "Печать таблицы");
                //                break;
                //            case 4:
                //                printDialog.PrintVisual(EmployeeDataGrid3, "Печать таблицы");
                //                break;
                //            case 5:
                //                printDialog.PrintVisual(EmployeeDataGrid4, "Печать таблицы");
                //                break;
                //            case 6:
                //                printDialog.PrintVisual(EmployeeDataGrid5, "Печать таблицы");
                //                break;
                //            case 7:
                //                printDialog.PrintVisual(EmployeeDataGrid6, "Печать таблицы");
                //                break;
                //            case 8:
                //                printDialog.PrintVisual(EmployeeDataGrid7, "Печать таблицы");
                //                break;
                //            case 9:
                //                printDialog.PrintVisual(EmployeeDataGrid8, "Печать таблицы");
                //                break;
                //            case 10:
                //                printDialog.PrintVisual(EmployeeDataGrid9, "Печать таблицы");
                //                break;
                //            case 11:
                //                printDialog.PrintVisual(EmployeeDataGrid10, "Печать таблицы");
                //                break;
                //            default:
                //                break;
                //        }
                //    }
                //}
                //catch(Exception)
                //{
                //    MessageBox.Show("Ошибка печати", "Ошибка операции", MessageBoxButton.OK, MessageBoxImage.Error);

                //}

            }

        private void SaveAsPdf_Click(object sender, RoutedEventArgs e) //сохранение datagrid как таблицы
        {
            APP = new Microsoft.Office.Interop.Excel.Application();
            OpenFileDialog folderBrowser = new OpenFileDialog();

            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            folderBrowser.FileName = "";
            try
            {
                Workbook book = SavePrintAction();
                PrintDialog dialogPrint = new PrintDialog();
                if (dialogPrint.ShowDialog() == true)
                {
                    book.PrintDocument.Print();
                }
                
            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка сохранения файла как pdf", MessageBoxButton.OK, MessageBoxImage.Error);
            }
                //try
                //{
                //    PrintDialog printDialog = new PrintDialog();
                //    MessageBox.Show("Сохранение как .pdf. Это может занять некоторое время, ожидайте", "Сохранение файла", MessageBoxButton.OK, MessageBoxImage.Information);

                //    switch (chosen_tab)
                //    {
                //        case 1:
                //            printDialog.PrintVisual(EmployeeDataGrid, "Сохранение таблицы");
                //            break;
                //        case 2:
                //            printDialog.PrintVisual(EmployeeDataGrid1, "Сохранение таблицы");
                //            break;
                //        case 4:
                //            printDialog.PrintVisual(EmployeeDataGrid3, "Сохранение таблицы");
                //            break;
                //        case 5:
                //            printDialog.PrintVisual(EmployeeDataGrid4, "Сохранение таблицы");
                //            break;
                //        case 6:
                //            printDialog.PrintVisual(EmployeeDataGrid5, "Сохранение таблицы");
                //            break;
                //        case 7:
                //            printDialog.PrintVisual(EmployeeDataGrid6, "Сохранение таблицы");
                //            break;
                //        case 8:
                //            printDialog.PrintVisual(EmployeeDataGrid7, "Сохранение таблицы");
                //            break;
                //        case 9:
                //            printDialog.PrintVisual(EmployeeDataGrid8, "Сохранение таблицы");
                //            break;
                //        case 10:
                //            printDialog.PrintVisual(EmployeeDataGrid9, "Сохранение таблицы");
                //            break;
                //        default:
                //            break;
                //    }
                //}
                //catch (Exception)
                //{
                //    MessageBox.Show("Ошибка сохранения", "Ошибка операции", MessageBoxButton.OK, MessageBoxImage.Error);
                //}

            }

        
        public void ChangeTabColor()
        {
            switch (chosen_tab)
            {
                case 1:
                    label1.Background = new SolidColorBrush(Color.FromRgb(222, 222, 222));
                    tabitem2.Background = new SolidColorBrush(Color.FromRgb(135, 141, 255));
                    tabitem4.Background = new SolidColorBrush(Color.FromRgb(255, 255, 80));
                    tabitem5.Background = new SolidColorBrush(Color.FromRgb(255, 82, 82));
                    tabitem6.Background = new SolidColorBrush(Color.FromRgb(194,229,253));
                    tabitem7.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem8.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem9.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem10.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem11.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));

                    break;
                case 2:
                    label1.Background = new SolidColorBrush(Color.FromRgb(217, 125, 232));
                    tabitem2.Background = new SolidColorBrush(Color.FromRgb(222, 222, 222));
                    tabitem4.Background = new SolidColorBrush(Color.FromRgb(255, 255, 80));
                    tabitem5.Background = new SolidColorBrush(Color.FromRgb(255, 82, 82));
                    tabitem6.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem7.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem8.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem9.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem10.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem11.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));

                    break;
               
                case 4:
                    label1.Background = new SolidColorBrush(Color.FromRgb(217, 125, 232));
                    tabitem2.Background = new SolidColorBrush(Color.FromRgb(135, 141, 255));
                    tabitem4.Background = new SolidColorBrush(Color.FromRgb(222, 222, 222));
                    tabitem5.Background = new SolidColorBrush(Color.FromRgb(255, 82, 82));
                    tabitem6.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem7.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem8.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem9.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem10.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem11.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));

                    break;
                case 5:
                    label1.Background = new SolidColorBrush(Color.FromRgb(217, 125, 232));
                    tabitem2.Background = new SolidColorBrush(Color.FromRgb(135, 141, 255));
                    tabitem4.Background = new SolidColorBrush(Color.FromRgb(255, 255, 80));
                    tabitem5.Background = new SolidColorBrush(Color.FromRgb(222, 222, 222));
                    tabitem6.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem7.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem8.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem9.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem10.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem11.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));

                    break;
                case 6:
                    label1.Background = new SolidColorBrush(Color.FromRgb(217, 125, 232));
                    tabitem2.Background = new SolidColorBrush(Color.FromRgb(135, 141, 255));
                    tabitem4.Background = new SolidColorBrush(Color.FromRgb(255, 255, 80));
                    tabitem5.Background = new SolidColorBrush(Color.FromRgb(255, 82, 82));
                    tabitem6.Background = new SolidColorBrush(Color.FromRgb(222,222,222));
                    tabitem7.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem8.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem9.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem10.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem11.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));

                    break;
                case 7:
                    label1.Background = new SolidColorBrush(Color.FromRgb(217, 125, 232));
                    tabitem2.Background = new SolidColorBrush(Color.FromRgb(135, 141, 255));
                    tabitem4.Background = new SolidColorBrush(Color.FromRgb(255, 255, 80));
                    tabitem5.Background = new SolidColorBrush(Color.FromRgb(255, 82, 82));
                    tabitem6.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem7.Background = new SolidColorBrush(Color.FromRgb(222, 222, 222));
                    tabitem8.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem9.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem10.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem11.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));

                    break;
                case 8:
                    label1.Background = new SolidColorBrush(Color.FromRgb(217, 125, 232));
                    tabitem2.Background = new SolidColorBrush(Color.FromRgb(135, 141, 255));
                    tabitem4.Background = new SolidColorBrush(Color.FromRgb(255, 255, 80));
                    tabitem5.Background = new SolidColorBrush(Color.FromRgb(255, 82, 82));
                    tabitem6.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem7.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem8.Background = new SolidColorBrush(Color.FromRgb(222, 222, 222));
                    tabitem9.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem10.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem11.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));

                    break;
                case 9:
                    label1.Background = new SolidColorBrush(Color.FromRgb(217, 125, 232));
                    tabitem2.Background = new SolidColorBrush(Color.FromRgb(135, 141, 255));
                    tabitem4.Background = new SolidColorBrush(Color.FromRgb(255, 255, 80));
                    tabitem5.Background = new SolidColorBrush(Color.FromRgb(255, 82, 82));
                    tabitem6.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem7.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem8.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem9.Background = new SolidColorBrush(Color.FromRgb(222, 222, 222));
                    tabitem10.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem11.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));

                    break;
                case 10:
                    label1.Background = new SolidColorBrush(Color.FromRgb(217, 125, 232));
                    tabitem2.Background = new SolidColorBrush(Color.FromRgb(135, 141, 255));
                    tabitem4.Background = new SolidColorBrush(Color.FromRgb(255, 255, 80));
                    tabitem5.Background = new SolidColorBrush(Color.FromRgb(255, 82, 82));
                    tabitem6.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem7.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem8.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem9.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem10.Background = new SolidColorBrush(Color.FromRgb(222, 222, 222));
                    tabitem11.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));

                    break;
                case 11:
                    label1.Background = new SolidColorBrush(Color.FromRgb(217, 125, 232));
                    tabitem2.Background = new SolidColorBrush(Color.FromRgb(135, 141, 255));
                    tabitem4.Background = new SolidColorBrush(Color.FromRgb(255, 255, 80));
                    tabitem5.Background = new SolidColorBrush(Color.FromRgb(255, 82, 82));
                    tabitem6.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem7.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem8.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem9.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));
                    tabitem11.Background = new SolidColorBrush(Color.FromRgb(222, 222, 222));
                    tabitem10.Background = new SolidColorBrush(Color.FromRgb(194, 229, 253));

                    break;
                default:
                    break;
            }
        }
        private void ChangeCellColor(object sender, RoutedEventArgs e)
        {
           
            DataGridCell cell = GetCell(selectedCell[0], selectedCell[1], EmployeeDataGrid);
            Color color = (Color)ColorConverter.ConvertFromString(sender.GetType().GetProperty("Background").GetValue(sender).ToString());
            //cell.Background = new SolidColorBrush(color);
            int row = selectedCell[0];
            int column = selectedCell[1];
            cellColorDictionary.Add(new int[2] { row, column }, color.ToString());
            isColorChanged[selectedCell[0]] = true;
           
        }
        //public void ChangeCellColor()
        //{

            
        //}
        public DataGridCell GetCell(int rowIndex, int columnIndex, DataGrid dg)
        {
            DataGridRow row = dg.ItemContainerGenerator.ContainerFromIndex(rowIndex) as DataGridRow;
            DataGridCellsPresenter p = GetVisualChild<DataGridCellsPresenter>(row);
            DataGridCell cell = p.ItemContainerGenerator.ContainerFromIndex(columnIndex) as DataGridCell;
            return cell;
        }

        private void EmployeeDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            //try
            //{
            int currentIndex = ((Employee)e.Row.Item).OrderId - 1;
            Console.WriteLine(currentIndex);
            Console.WriteLine(isColorChanged[currentIndex]);
            foreach (var item in cellColorDictionary)
            {
                Console.WriteLine(item.Key[0]);
            }
            try
            {
                if (isColorChanged[currentIndex])
                {
                    Color color = (Color)ColorConverter.ConvertFromString(cellColorDictionary.FirstOrDefault(x => x.Key[0] == currentIndex).Value);
                    DataGridRow row = EmployeeDataGrid.ItemContainerGenerator.ContainerFromIndex(cellColorDictionary.FirstOrDefault(x => x.Key[0] == currentIndex).Key[0]) as DataGridRow;
                    Console.WriteLine(((Employee)row.Item).Name);
                    DataGridCellsPresenter p = GetVisualChild<DataGridCellsPresenter>(row);
                    DataGridCell cell = p.ItemContainerGenerator.ContainerFromIndex(cellColorDictionary.FirstOrDefault(x => x.Key[0] == currentIndex).Key[1]) as DataGridCell;
                    cell.Background = new SolidColorBrush(color);
                }
            }
            catch (Exception e8)
            {
                MessageBox.Show(e8.Message + e8.StackTrace);
            }




            //SolidColorBrush nb = new SolidColorBrush(color);
            //e.Row.Background = nb;


            //}
            //catch(Exception)
            //{
            //    e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            //}

        }
        //private void EmployeeDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        //{
        //    try
        //    {
        //        Color color = (Color)ColorConverter.ConvertFromString(cellColorList[((Employee)e.Row.Item).OrderId - 1]);
        //        SolidColorBrush nb = new SolidColorBrush(color);
        //        e.Row.Background = nb;

        //    }
        //    catch (Exception)
        //    {
        //        e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
        //    }
        //}

        private void EmployeeDataGrid_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            DependencyObject dep = (DependencyObject)e.OriginalSource;

            while ((dep != null) && !(dep is DataGridCell) && !(dep is DataGridColumnHeader))
            {
                dep = VisualTreeHelper.GetParent(dep);
            }

            if (dep == null)
                return;

            if (dep is DataGridColumnHeader)
            {
                DataGridColumnHeader columnHeader = dep as DataGridColumnHeader;

                // find the property that this cell's column is bound to
                string boundPropertyName = FindBoundProperty(columnHeader.Column);

                int columnIndex = columnHeader.Column.DisplayIndex;

            }

            if (dep is DataGridCell)
            {
                DataGridCell cell = dep as DataGridCell;

                // navigate further up the tree
                while ((dep != null) && !(dep is DataGridRow))
                {
                    dep = VisualTreeHelper.GetParent(dep);
                }

                if (dep == null)
                    return;

                DataGridRow row = dep as DataGridRow;

                object value = ExtractBoundValue(row, cell);

                int columnIndex = cell.Column.DisplayIndex;
                int rowIndex = FindRowIndex(row);
                selectedCell[0] = rowIndex;
                selectedCell[1] = columnIndex;
                
                
            }
        }
        /// <summary>
        /// Determine the index of a DataGridRow
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private int FindRowIndex(DataGridRow row)
        {
            DataGrid dataGrid = ItemsControl.ItemsControlFromItemContainer(row) as DataGrid;

            int index = dataGrid.ItemContainerGenerator.IndexFromContainer(row);

            return index;
        }

        /// <summary>
        /// Find the value that is bound to a DataGridCell
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        private object ExtractBoundValue(DataGridRow row, DataGridCell cell)
        {
            // find the property that this cell's column is bound to
            string boundPropertyName = FindBoundProperty(cell.Column);

            // find the object that is realted to this row
            object data = row.Item;

            // extract the property value
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(data);
            PropertyDescriptor property = properties[boundPropertyName];
            object value = property.GetValue(data);

            return value;
        }

        /// <summary>
        /// Find the name of the property which is bound to the given column
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        private string FindBoundProperty(DataGridColumn col)
        {
            DataGridBoundColumn boundColumn = col as DataGridBoundColumn;

            // find the property that this column is bound to
            Binding binding = boundColumn.Binding as Binding;
            string boundPropertyName = binding.Path.Path;

            return boundPropertyName;
        }


        private void EmployeeDataGrid1_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            try
            {
                Color color = (Color)ColorConverter.ConvertFromString(cellColorList[((Employee)e.Row.Item).OrderId - 1]);
                SolidColorBrush nb = new SolidColorBrush(color);
                e.Row.Background = nb;

            }
            catch (Exception)
            {
                e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            }
        }

        private void EmployeeDataGrid3_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            try
            {
                Color color = (Color)ColorConverter.ConvertFromString(cellColorList[((Employee)e.Row.Item).OrderId - 1]);
                SolidColorBrush nb = new SolidColorBrush(color);
                e.Row.Background = nb;

            }
            catch (Exception)
            {
                e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            }
        }

        private void EmployeeDataGrid4_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            try
            {
                Color color = (Color)ColorConverter.ConvertFromString(cellColorList[((Employee)e.Row.Item).OrderId - 1]);
                SolidColorBrush nb = new SolidColorBrush(color);
                e.Row.Background = nb;

            }
            catch (Exception)
            {
                e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            }
        }

        private void EmployeeDataGrid5_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            try
            {
                Color color = (Color)ColorConverter.ConvertFromString(cellColorList[((Employee)e.Row.Item).OrderId - 1]);
                SolidColorBrush nb = new SolidColorBrush(color);
                e.Row.Background = nb;

            }
            catch (Exception)
            {
                e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            }
        }

        private void EmployeeDataGrid10_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            try
            {
                Color color = (Color)ColorConverter.ConvertFromString(cellColorList[((Employee)e.Row.Item).OrderId - 1]);
                SolidColorBrush nb = new SolidColorBrush(color);
                e.Row.Background = nb;

            }
            catch (Exception)
            {
                e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            }
        }

        private void EmployeeDataGrid6_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            try
            {
                Color color = (Color)ColorConverter.ConvertFromString(cellColorList[((Employee)e.Row.Item).OrderId - 1]);
                SolidColorBrush nb = new SolidColorBrush(color);
                e.Row.Background = nb;

            }
            catch (Exception)
            {
                e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            }
        }

        private void EmployeeDataGrid7_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            try
            {
                Color color = (Color)ColorConverter.ConvertFromString(cellColorList[((Employee)e.Row.Item).OrderId - 1]);
                SolidColorBrush nb = new SolidColorBrush(color);
                e.Row.Background = nb;

            }
            catch (Exception)
            {
                e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            }
        }

        private void EmployeeDataGrid8_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            try
            {
                Color color = (Color)ColorConverter.ConvertFromString(cellColorList[((Employee)e.Row.Item).OrderId - 1]);
                SolidColorBrush nb = new SolidColorBrush(color);
                e.Row.Background = nb;

            }
            catch (Exception)
            {
                e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            }
        }

        private void EmployeeDataGrid9_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            try
            {
                Color color = (Color)ColorConverter.ConvertFromString(cellColorList[((Employee)e.Row.Item).OrderId - 1]);
                SolidColorBrush nb = new SolidColorBrush(color);
                e.Row.Background = nb;

            }
            catch (Exception)
            {
                e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            }
        }

        private void ColorBlue(object sender, RoutedEventArgs e)
        {
            switch(chosen_tab)
            {
                case 1:
                    ColorGridRow(EmployeeDataGrid, "#FFC2E5FD");
                    break;
                case 2:
                    ColorGridRow(EmployeeDataGrid1, "#FFC2E5FD");
                    break;
                case 4:
                    ColorGridRow(EmployeeDataGrid3, "#FFC2E5FD");
                    break;
                case 5:
                    ColorGridRow(EmployeeDataGrid4, "#FFC2E5FD");
                    break;
                case 6:
                    ColorGridRow(EmployeeDataGrid5, "#FFC2E5FD");
                    break;
                case 7:
                    ColorGridRow(EmployeeDataGrid6, "#FFC2E5FD");
                    break;
                case 8:
                    ColorGridRow(EmployeeDataGrid7, "#FFC2E5FD");
                    break;
                case 9:
                    ColorGridRow(EmployeeDataGrid8, "#FFC2E5FD");
                    break;
                case 10:
                    ColorGridRow(EmployeeDataGrid9, "#FFC2E5FD");
                    break;
                case 11:
                    ColorGridRow(EmployeeDataGrid10, "#FFC2E5FD");
                    break;

                    
            }
        }

        private void ColorGreen(object sender, RoutedEventArgs e)
        {
            switch (chosen_tab)
            {
                case 1:
                    ColorGridRow(EmployeeDataGrid, "#FFAFF99E");
                    break;
                case 2:
                    ColorGridRow(EmployeeDataGrid1, "#FFAFF99E");
                    break;
                case 4:
                    ColorGridRow(EmployeeDataGrid3, "#FFAFF99E");
                    break;
                case 5:
                    ColorGridRow(EmployeeDataGrid4, "#FFAFF99E");
                    break;
                case 6:
                    ColorGridRow(EmployeeDataGrid5, "#FFAFF99E");
                    break;
                case 7:
                    ColorGridRow(EmployeeDataGrid6, "#FFAFF99E");
                    break;
                case 8:
                    ColorGridRow(EmployeeDataGrid7, "#FFAFF99E");
                    break;
                case 9:
                    ColorGridRow(EmployeeDataGrid8, "#FFAFF99E");
                    break;
                case 10:
                    ColorGridRow(EmployeeDataGrid9, "#FFAFF99E");
                    break;
                case 11:
                    ColorGridRow(EmployeeDataGrid10, "#FFAFF99E");
                    break;


            }
        }

        private void ColorYellow(object sender, RoutedEventArgs e)
        {
            switch (chosen_tab)
            {
                case 1:
                    ColorGridRow(EmployeeDataGrid, "#e9ff59");
                    break;
                case 2:
                    ColorGridRow(EmployeeDataGrid1, "#e9ff59");
                    break;
                case 4:
                    ColorGridRow(EmployeeDataGrid3, "#e9ff59");
                    break;
                case 5:
                    ColorGridRow(EmployeeDataGrid4, "#e9ff59");
                    break;
                case 6:
                    ColorGridRow(EmployeeDataGrid5, "#e9ff59");
                    break;
                case 7:
                    ColorGridRow(EmployeeDataGrid6, "#e9ff59");
                    break;
                case 8:
                    ColorGridRow(EmployeeDataGrid7, "#e9ff59");
                    break;
                case 9:
                    ColorGridRow(EmployeeDataGrid8, "#e9ff59");
                    break;
                case 10:
                    ColorGridRow(EmployeeDataGrid9, "#e9ff59");
                    break;
                case 11:
                    ColorGridRow(EmployeeDataGrid10, "#e9ff59");
                    break;


            }
        }

        private void ColorRed(object sender, RoutedEventArgs e)
        {
            switch (chosen_tab)
            {
                case 1:
                    ColorGridRow(EmployeeDataGrid, "#FFFFBE71");
                    break;
                case 2:
                    ColorGridRow(EmployeeDataGrid1, "#FFFFBE71");
                    break;
                case 4:
                    ColorGridRow(EmployeeDataGrid3, "#FFFFBE71");
                    break;
                case 5:
                    ColorGridRow(EmployeeDataGrid4, "#FFFFBE71");
                    break;
                case 6:
                    ColorGridRow(EmployeeDataGrid5, "#FFFFBE71");
                    break;
                case 7:
                    ColorGridRow(EmployeeDataGrid6, "#FFFFBE71");
                    break;
                case 8:
                    ColorGridRow(EmployeeDataGrid7, "#FFFFBE71");
                    break;
                case 9:
                    ColorGridRow(EmployeeDataGrid8, "#FFFFBE71");
                    break;
                case 10:
                    ColorGridRow(EmployeeDataGrid9, "#FFFFBE71");
                    break;
                case 11:
                    ColorGridRow(EmployeeDataGrid10, "#FFFFBE71");
                    break;


            }
        }

        private void ColorGridRow(DataGrid grid, string color)
        {
            Employee employee = (Employee)grid.SelectedItems[0];
            string command = "UPDATE Employees SET cellColor='" + color + "' WHERE OrderId='" + employee.OrderId + "'";
            db.Database.ExecuteSqlCommand(command);
            cellColorList[employee.OrderId-1] = color;
        }

        private void ColorWhite(object sender, RoutedEventArgs e)
        {
            switch (chosen_tab)
            {
                case 1:
                    ColorGridRow(EmployeeDataGrid, "#FFFFFFFF");
                    break;
                case 2:
                    ColorGridRow(EmployeeDataGrid1, "#FFFFFFFF");
                    break;
                case 4:
                    ColorGridRow(EmployeeDataGrid3, "#FFFFFFFF");
                    break;
                case 5:
                    ColorGridRow(EmployeeDataGrid4, "#FFFFFFFF");
                    break;
                case 6:
                    ColorGridRow(EmployeeDataGrid5, "#FFFFFFFF");
                    break;
                case 7:
                    ColorGridRow(EmployeeDataGrid6, "#FFFFFFFF");
                    break;
                case 8:
                    ColorGridRow(EmployeeDataGrid7, "#FFFFFFFF");
                    break;
                case 9:
                    ColorGridRow(EmployeeDataGrid8, "#FFFFFFFF");
                    break;
                case 10:
                    ColorGridRow(EmployeeDataGrid9, "#FFFFFFFF");
                    break;
                case 11:
                    ColorGridRow(EmployeeDataGrid10, "#FFFFFFFF");
                    break;


            }
        }
        private void EmployeeDataGrid_CopyingRowClipboardContent(object sender, DataGridRowClipboardEventArgs e)
        {
            if (clickedContextMenu==0)
            {
                var currentCell = e.ClipboardRowContent[EmployeeDataGrid.CurrentCell.Column.DisplayIndex];
                e.ClipboardRowContent.Clear();
                e.ClipboardRowContent.Add(currentCell);
            }
            
        }

        private void EmployeeDataGrid1_CopyingRowClipboardContent(object sender, DataGridRowClipboardEventArgs e)
        {
            if (clickedContextMenu == 0)
            {
                var currentCell = e.ClipboardRowContent[EmployeeDataGrid1.CurrentCell.Column.DisplayIndex];
                e.ClipboardRowContent.Clear();
                e.ClipboardRowContent.Add(currentCell);
            }

        }

        private void EmployeeDataGrid3_CopyingRowClipboardContent(object sender, DataGridRowClipboardEventArgs e)
        {

            if (clickedContextMenu == 0)
            {
                var currentCell = e.ClipboardRowContent[EmployeeDataGrid3.CurrentCell.Column.DisplayIndex];
                e.ClipboardRowContent.Clear();
                e.ClipboardRowContent.Add(currentCell);
            }

        }

        private void EmployeeDataGrid4_CopyingRowClipboardContent(object sender, DataGridRowClipboardEventArgs e)
        {

            if (clickedContextMenu == 0)
            {
                var currentCell = e.ClipboardRowContent[EmployeeDataGrid4.CurrentCell.Column.DisplayIndex];
                e.ClipboardRowContent.Clear();
                e.ClipboardRowContent.Add(currentCell);
            }

        }

        private void EmployeeDataGrid5_CopyingRowClipboardContent(object sender, DataGridRowClipboardEventArgs e)
        {
            if (clickedContextMenu == 0)
            {
                var currentCell = e.ClipboardRowContent[EmployeeDataGrid5.CurrentCell.Column.DisplayIndex];
                e.ClipboardRowContent.Clear();
                e.ClipboardRowContent.Add(currentCell);
            }

        }

        private void EmployeeDataGrid10_CopyingRowClipboardContent(object sender, DataGridRowClipboardEventArgs e)
        {
            if (clickedContextMenu == 0)
            {
                var currentCell = e.ClipboardRowContent[EmployeeDataGrid10.CurrentCell.Column.DisplayIndex];
                e.ClipboardRowContent.Clear();
                e.ClipboardRowContent.Add(currentCell);
            }

        }

        private void EmployeeDataGrid6_CopyingRowClipboardContent(object sender, DataGridRowClipboardEventArgs e)
        {
            if (clickedContextMenu == 0)
            {
                var currentCell = e.ClipboardRowContent[EmployeeDataGrid6.CurrentCell.Column.DisplayIndex];
                e.ClipboardRowContent.Clear();
                e.ClipboardRowContent.Add(currentCell);
            }

        }

        private void EmployeeDataGrid7_CopyingRowClipboardContent(object sender, DataGridRowClipboardEventArgs e)
        {
            if (clickedContextMenu == 0)
            {
                var currentCell = e.ClipboardRowContent[EmployeeDataGrid7.CurrentCell.Column.DisplayIndex];
                e.ClipboardRowContent.Clear();
                e.ClipboardRowContent.Add(currentCell);
            }

        }
        private void EmployeeDataGrid8_CopyingRowClipboardContent(object sender, DataGridRowClipboardEventArgs e)
        {
            if (clickedContextMenu == 0)
            {
                var currentCell = e.ClipboardRowContent[EmployeeDataGrid8.CurrentCell.Column.DisplayIndex];
                e.ClipboardRowContent.Clear();
                e.ClipboardRowContent.Add(currentCell);
            }
        }

        private void EmployeeDataGrid9_CopyingRowClipboardContent(object sender, DataGridRowClipboardEventArgs e)
        {
            if (clickedContextMenu == 0)
            {
                var currentCell = e.ClipboardRowContent[EmployeeDataGrid9.CurrentCell.Column.DisplayIndex];
                e.ClipboardRowContent.Clear();
                e.ClipboardRowContent.Add(currentCell);
            }
        }

        private void EmployeeDataGrid_KeyDown(object sender, KeyEventArgs e)
        {

            //if (e.Key == Key.V && Keyboard.Modifiers == ModifierKeys.Control)
            //{
            //    string data = Clipboard.GetData(DataFormats.Text).ToString();
            //    string[] cells = data.Split('\t');
            //    for (int i = 0; i < cells.Length; i++)
            //        Console.WriteLine(cells[i]);
            //        //employeedatagrid[employeedatagrid.selecteditems[0], i] = cells[i];
            //}
        }


        private void CopyClick(object sender, RoutedEventArgs e)
        {
            clickedContextMenu = 0;
        }
        private void CopyClick1(object sender, RoutedEventArgs e)
        {
            clickedContextMenu = 1;
        }

        private void EmployeeDataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            int b = e.Row.GetIndex();
            BeginEdit(EmployeeDataGrid, b);

        }
        public void BeginEdit(DataGrid grid, int b)
        {
            DataGridRow row = (DataGridRow)grid.ItemContainerGenerator.ContainerFromIndex(b);
            Employee employee = (Employee)row.Item;
            Employee employee2 = new Employee(
                employee.OrderId, employee.Name, employee.Position, employee.Department, employee.ExaminationDateFact,
                employee.ExaminationDatePlan, employee.ExaminationComplexDateFact, employee.ExaminationComplexDatePlan,
                employee.AttestationDateFact, employee.AttestationDatePlan, employee.PbminimumPassDateFact, employee.PbminimumPassDatePlan,
                employee.MedicalCheckDateFact, employee.MedicalCheckDatePlan, Convert.ToInt32(employee.TabNumber), employee.BirthDate, employee.EntryDate,
                employee.RelocationDate, employee.PrimaryInstructionDate, employee.InternshipDate, employee.InternshipDetails,
                employee.DublicationDate, employee.DublicationDetails, employee.IndependentDate, employee.IndependentDetails, employee.ExtraStatus,
                employee.CellColor);
            operationsDoneDetails.Push(employee2);
            operationsDone.Push("update");
        }
        private void EmployeeDataGrid3_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            int b = e.Row.GetIndex();
            BeginEdit(EmployeeDataGrid3, b);
        }

        private void EmployeeDataGrid4_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            int b = e.Row.GetIndex();
            BeginEdit(EmployeeDataGrid4, b);
        }

        private void EmployeeDataGrid5_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            int b = e.Row.GetIndex();
            BeginEdit(EmployeeDataGrid5, b);
        }

        private void EmployeeDataGrid10_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            int b = e.Row.GetIndex();
            BeginEdit(EmployeeDataGrid10, b);
        }

        private void EmployeeDataGrid6_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            int b = e.Row.GetIndex();
            BeginEdit(EmployeeDataGrid6, b);
        }

        private void EmployeeDataGrid7_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            int b = e.Row.GetIndex();
            BeginEdit(EmployeeDataGrid7, b);
        }

        private void EmployeeDataGrid8_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            int b = e.Row.GetIndex();
            BeginEdit(EmployeeDataGrid8, b);
        }

        private void EmployeeDataGrid9_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            int b = e.Row.GetIndex();
            BeginEdit(EmployeeDataGrid9, b);
        }

        private void EmployeeDataGrid1_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            int b = e.Row.GetIndex();
            BeginEdit(EmployeeDataGrid1, b);
        }

        
        static T GetVisualChild<T>(Visual parent) where T : Visual
        {
            T child = default(T);
            int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < numVisuals; i++)
            {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent, i);
                child = v as T;
                if (child == null)
                {
                    child = GetVisualChild<T>(v);
                }
                if (child != null)
                {
                    break;
                }
            }
            return child;
        }
    }
}


