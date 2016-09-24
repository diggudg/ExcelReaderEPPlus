using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace ExcelReaderUsingOpenOfficeXML.Models
{
    public class Student
    {
        [Key]
        public int StudentId { get; set; }

        public int? ClientId { get; set; }

        [StringLength(45)]
        public string StudName { get; set; }

        [StringLength(45)]
        public string StudLastName { get; set; }

        [StringLength(45)]
        public string StudMiddleName { get; set; }

        [StringLength(100)]
        public string StudFathername { get; set; }

        [StringLength(100)]
        public string StudMotherName { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public DateTime? CreationDate { get; set; }

        [Column(TypeName = "date")]
        public DateTime? DOB { get; set; }
    }


    public class StudentPersonalDetails
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public string FatherName { get; set; }
        public string MotherName { get; set; }
        public string DOB { get; set; }




    }
    public class StudentAddress
    {
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public int PIN { get; set; }
    }
    public class StudentContacts
    {
        public string Email { get; set; }
        public string AlternatePhone { get; set; }
        public string Phone { get; set; }
    }
    public class StudentCourseDetails
    {
        public int StudentId { get; set; }
        public string CourseName { get; set; }
        public string CourseDetails { get; set; }

    }
    public class CourseDetails
    {
        public int CourseID { get; set; }
        public string CourseName { get; set; }
        public string CourseDetail { get; set; }
    }
    public class StudentsDetails
    {
        public StudentsDetails()
        {
            StudPersonalDetails = new StudentPersonalDetails();
            Clients = new Clients();
            StudentAddress = new StudentAddress();
            CourseDetails = new CourseDetails();
            StudentContacts = new StudentContacts();
        }
        public StudentPersonalDetails StudPersonalDetails { get; set; }
        public Clients Clients { get; set; }
        public StudentAddress StudentAddress { get; set; }
        public CourseDetails CourseDetails { get; set; }
        public StudentContacts StudentContacts { get; set; }
    }
    public class Clients
    {
        public int ClientID { get; set; }
        public string ClientName { get; set; }
        public string ClientDescriptions { get; set; }
    }

}