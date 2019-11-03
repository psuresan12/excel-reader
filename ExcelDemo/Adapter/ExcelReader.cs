using Ganss.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelDemo.Adapter
{
    public class ExcelReader
    {
        public ExcelReader()
        {
            
        }

        public IEnumerable<Person> Read()
        {
            String file_name = "C:\\Users\\prataps\\source\\repos\\TestExcel.xlsx";
            //FileStream file_stream = new FileStream(file_name, FileMode.Create);

            var excel = new ExcelMapper() { TrackObjects = true };
           
            var persons = excel.Fetch<Person>(file_name, "Customers").ToList();
            persons[3].Gender = "Female";

            //saving first time
            String fileOut = "C:\\Users\\prataps\\source\\repos\\TestExcel_out.xlsx";
            excel.Save(fileOut, "Customers");

            //change the values
            persons[1].Gender = "Unknown";
            //saving for the second time
            excel.Save(fileOut, "Customers");



            //changing the value for third time
            persons[2].Gender = "Dont Care";

            //saving for the third time
            excel.Save(fileOut, persons);





            return persons;
        }


        public IEnumerable<Person> ReadStream()
        {
            String file = "C:\\Users\\prataps\\source\\repos\\TestExcel.xlsx";
            List<Person> persons;
            using (FileStream readStream = new FileStream(file, FileMode.Open,FileAccess.Read))
            {
                var excel = new ExcelMapper() { TrackObjects = true };
                persons = excel.Fetch<Person>(readStream, "Customers").ToList();
                readStream.Close();
            }

            return persons;
        }


        public IEnumerable<Person> updateStream()
        {
            String file = "C:\\Users\\prataps\\source\\repos\\TestExcel.xlsx";
            List<Person> persons = ReadStream().ToList();
            if (persons == null)
                persons = new List<Person>();

            persons.Add(new Person() { Name = "Rahul"+(persons.Count+1), Gender = "Male" });

           
            using (FileStream writeStream = new FileStream(file, FileMode.Open, FileAccess.ReadWrite))
            {
                var excel = new ExcelMapper() { TrackObjects = true };
                excel.Save(writeStream, persons, "Customers", false);
                writeStream.Close();

            }
            return persons;
        }

        public IEnumerable<Person> updateFile()
        {
            String file = "C:\\Users\\prataps\\source\\repos\\TestExcel.xlsx";
            List<Person> persons = ReadStream().ToList();
            if (persons == null)
                persons = new List<Person>();

            persons.Add(new Person() { Name = "Rahul" + (persons.Count + 1), Gender = "Male" });

            var excel = new ExcelMapper() { TrackObjects = true };
            excel.Save(file, persons, "Customers");
            
            return persons;
        }


    }


    public class Person
    {
       // public int Id { get; set; }
        public string Name { get; set; }
        //public string Email { get; set; }
       

        public string Gender { get; set; }
    }

    public class Post
    {
        public int PostId { get; set; }
        public DateTime CreatedDate { get; set; }
    }



}
