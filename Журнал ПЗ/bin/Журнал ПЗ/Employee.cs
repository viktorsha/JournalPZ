using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace График_ПЗ
{
    public class Employee
    {
        private int id;
        private int orderId;
        private string name;
        private string position;
        private string department;
        private string examinationDateFact;
        private string examinationDatePlan;
        private string examinationComplexDatePlan;
        private string examinationComplexDateFact;
        private string attestationDateFact;
        private string attestationDatePlan;
        private string pbminimumPassDateFact;
        private string pbminimumPassDatePlan;
        private string medicalCheckDateFact;
        private string medicalCheckDatePlan;
        private long? tabNumber;
        private string birthDate;
        private string entryDate;
        private string relocationDate;
        private string primaryInstructionDate;
        private string internshipDate;
        private string internshipDetails;
        private string dublicationDate;
        private string dublicationDetails;
        private string independentDate;
        private string independentDetails;
        private string extraStatus;
        private string cellColor;


        public int Id
        {
            get { return id; }
            set { id = value; }
        }
        public int OrderId
        {
            get { return orderId; }
            set { orderId = value; }
        }
        
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        
        public string Position
        {
            get { return position; }
            set { position = value; }
        }
        public string Department
        {
            get { return department; }
            set { department = value; }
        }
        public string ExaminationDateFact
        {
            get { return examinationDateFact; }
            set { examinationDateFact = value; }
        }
        public string ExaminationDatePlan
        {
            get { return examinationDatePlan; }
            set { examinationDatePlan = value; }
        }
        public string ExaminationComplexDatePlan
        {
            get { return examinationComplexDatePlan; }
            set { examinationComplexDatePlan = value; }
        }
        public string ExaminationComplexDateFact
        {
            get { return examinationComplexDateFact; }
            set { examinationComplexDateFact = value; }
        }
        public string AttestationDateFact
        {
            get { return attestationDateFact; }
            set { attestationDateFact = value; }
        }
        public string AttestationDatePlan
        {
            get { return attestationDatePlan; }
            set { attestationDatePlan = value; }
        }
        public string PbminimumPassDateFact
        {
            get { return pbminimumPassDateFact; }
            set { pbminimumPassDateFact = value; }
        }
        public string PbminimumPassDatePlan
        {
            get { return pbminimumPassDatePlan; }
            set { pbminimumPassDatePlan = value; }
        }
        public string MedicalCheckDateFact
        {
            get { return medicalCheckDateFact; }
            set { medicalCheckDateFact = value; }
        }
        public string MedicalCheckDatePlan
        {
            get { return medicalCheckDatePlan; }
            set { medicalCheckDatePlan = value; }
        }
        public long? TabNumber
        {
            get { return tabNumber; }
            set { tabNumber = Convert.ToInt32(value); }
        }
        public string BirthDate
        {
            get { return birthDate; }
            set { birthDate = value; }
        }
        public string EntryDate
        {
            get { return entryDate; }
            set { entryDate = value; }
        }
        public string RelocationDate
        {
            get { return relocationDate; }
            set { relocationDate = value; }
        }
        public string PrimaryInstructionDate
        {
            get { return primaryInstructionDate; }
            set { primaryInstructionDate = value; }
        }
        public string InternshipDate
        {
            get { return internshipDate; }
            set { internshipDate = value; }
        }
        public string InternshipDetails
        {
            get { return internshipDetails; }
            set { internshipDetails = value; }
        }
        public string DublicationDate
        {
            get { return dublicationDate; }
            set { dublicationDate = value; }
        }
        public string DublicationDetails
        {
            get { return dublicationDetails; }
            set { dublicationDetails = value; }
        }
        public string IndependentDate
        {
            get { return independentDate; }
            set { independentDate = value; }
        }
        public string IndependentDetails
        {
            get { return independentDetails; }
            set { independentDetails = value; }
        }
        public string ExtraStatus
        {
            get { return extraStatus; }
            set { extraStatus = value; }
        }

        public string CellColor
        {
            get { return cellColor; }
            set { cellColor = value; }
        }
        public Employee()
        {

        }

        public Employee(int orderId, string name, string position, string department, string examinationDateFact, string examinationDatePlan, string examinationComplexDatePlan, string examinationComplexDateFact, string attestationDateFact, string attestationDatePlan, string pbminimumPassDateFact, string pbminimumPassDatePlan, string medicalCheckDateFact, string medicalCheckDatePlan, int tabNumber, string birthDate, string entryDate, string relocationDate, string primaryInstructionDate, string internshipDate, string internshipDetails, string dublicationDate, string dublicationDetails, string independentDate, string independentDetails, string extraStatus, string cellColor)
        {
            this.orderId = orderId;
            this.name = name;
            this.position = position;
            this.department = department;
            this.examinationDateFact = examinationDateFact;
            this.examinationDatePlan = examinationDatePlan;
            this.examinationComplexDatePlan = examinationComplexDatePlan;
            this.examinationComplexDateFact = examinationComplexDateFact;
            this.attestationDateFact = attestationDateFact;
            this.attestationDatePlan = attestationDatePlan;
            this.pbminimumPassDateFact = pbminimumPassDateFact;
            this.pbminimumPassDatePlan = pbminimumPassDatePlan;
            this.medicalCheckDateFact = medicalCheckDateFact;
            this.medicalCheckDatePlan = medicalCheckDatePlan;
            this.tabNumber = tabNumber;
            this.birthDate = birthDate;
            this.entryDate = entryDate;
            this.relocationDate = relocationDate;
            this.primaryInstructionDate = primaryInstructionDate;
            this.internshipDate = internshipDate;
            this.internshipDetails = internshipDetails;
            this.dublicationDate = dublicationDate;
            this.dublicationDetails = dublicationDetails;
            this.independentDate = independentDate;
            this.independentDetails = independentDetails;
            this.extraStatus = extraStatus;
            this.cellColor = cellColor;
        }
    }

}
