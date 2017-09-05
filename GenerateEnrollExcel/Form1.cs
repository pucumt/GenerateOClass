using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using MongoDB.Bson;
using MongoDB.Driver;
using NPOI.XSSF.UserModel;
using System.Data.OleDb;
using System.Threading;

namespace GenerateEnrollExcel
{
    public partial class Frm_Generator : Form
    {
        public Frm_Generator()
        {
            InitializeComponent();
        }

        private void btn_select_Click(object sender, EventArgs e)
        {
            //var result = ofD_file.ShowDialog();
            //txtFile.Text = ofD_file.FileName;
        }

        private void btn_generator_Click(object sender, EventArgs e)
        {
            var dbHelper = new DBHelper();
            var courses = dbHelper.getAllCourses();

            foreach (var curClass in courses)
            {
                BsonValue teacherId;
                if (!curClass.TryGetValue("teacherId", out teacherId))
                {
                    continue;
                }

                XSSFWorkbook hssfworkbook = null;
                using (FileStream fs = File.Open(@"template.xlsx", FileMode.Open,
                FileAccess.Read, FileShare.ReadWrite))
                {
                    //把xls文件读入workbook变量里，之后就可以关闭了  
                    hssfworkbook = new XSSFWorkbook(fs);
                    fs.Close();
                }

                var fileName = string.Format("沟通_{0}_{1}_{2}_{3}.xlsx", curClass.GetValue("schoolArea"), curClass.GetValue("subjectName"), curClass.GetValue("gradeName"), curClass.GetValue("name"));

                XSSFSheet sheet1 = hssfworkbook.GetSheet("template") as XSSFSheet;
                hssfworkbook.SetSheetName(0, curClass.GetValue("name").ToString());

                var curSource = dbHelper.getAllSources(curClass.GetValue("_id").ToString());
                var curTeacher = dbHelper.getTeacher(teacherId.ToString());
                
                for (var i = 0; i < curSource.Count; i++)
                {
                    var curStudent = dbHelper.getStudent(curSource[i].GetValue("studentId").ToString());

                    sheet1.GetRow(i + 2 + 6 * i).GetCell(1).SetCellValue(curClass.GetValue("name").ToString());
                    sheet1.GetRow(i + 2 + 6 * i).GetCell(2).SetCellValue(curStudent.GetValue("name").ToString());
                    sheet1.GetRow(i + 2 + 6 * i).GetCell(3).SetCellValue(curClass.GetValue("teacherName").ToString());
                    sheet1.GetRow(i + 2 + 6 * i).GetCell(4).SetCellValue(curStudent.GetValue("mobile").ToString());
                    sheet1.GetRow(i + 2 + 6 * i).GetCell(5).SetCellValue(curTeacher.GetValue("mobile").ToString());
                }

                sheet1.ForceFormulaRecalculation = true;

                using (FileStream fileStream = File.Open(fileName,
                    FileMode.Create, FileAccess.ReadWrite))
                {
                    hssfworkbook.Write(fileStream);
                    fileStream.Close();
                }

            }
            MessageBox.Show("导出成功！");
        }

        public List<Order> ExcelToDS(string Path)
        {
            FileStream file = new FileStream(Path, FileMode.Open, FileAccess.Read);
            XSSFWorkbook hssfworkbook = new XSSFWorkbook(file);
            XSSFSheet sheet1 = hssfworkbook.GetSheet("报名情况") as XSSFSheet;
            List<Order> orders = new List<Order>();
            var i = 1;
            while(sheet1.GetRow(i)!=null)
            {
                var newOrder = new Order();
                for (var j = 0; j < 11; j++)
                {
                    newOrder.studentName = sheet1.GetRow(i).GetCell(0).StringCellValue;
                    newOrder.mobile = sheet1.GetRow(i).GetCell(1).StringCellValue;
                    newOrder.className = sheet1.GetRow(i).GetCell(8).StringCellValue;
                    newOrder.school = sheet1.GetRow(i).GetCell(7).StringCellValue;
                    newOrder.grade = sheet1.GetRow(i).GetCell(10).StringCellValue;
                    newOrder.subject = sheet1.GetRow(i).GetCell(6).StringCellValue;
                    newOrder.courseTime = sheet1.GetRow(i).GetCell(9).StringCellValue;
                    newOrder.studentSchool = (sheet1.GetRow(i).GetCell(2)!=null?sheet1.GetRow(i).GetCell(2).StringCellValue:"");
                    newOrder.studentClass = (sheet1.GetRow(i).GetCell(3)!=null?sheet1.GetRow(i).GetCell(3).StringCellValue:"");
                }
                if(newOrder.className!="")
                {
                    orders.Add(newOrder);
                }                
                i++;
            }

             return orders;
        }
    }

    public class Order
    {
        public string studentName;
        public string mobile;
        public string studentSchool;
        public string studentClass;
        public string sex;
        public string subject;
        public string school;
        public string className;
        public string grade;
        public string courseTime;
    }

    public class Course
    {
        public string _Id;
        public string name;
        public string schoolArea;
        public string subjectName;
        public string gradeName;
    }

    public class DBHelper
    {
        //连接信息
        string conn = "mongodb://127.0.0.1:10105";
        string database = "websiteOnline2";
        MongoClient mongodb;
        IMongoDatabase mongoDataBase;


        public DBHelper() {
            mongodb = new MongoClient(conn);//连接数据库
            mongoDataBase = mongodb.GetDatabase(database);//选择数据库名
        }

        public List<BsonDocument> getAllCourses()
        {
            IMongoCollection<BsonDocument> mongoCollection = mongoDataBase.GetCollection <BsonDocument> ("trainClasss");//选择集合，相当于表

            //var count = mongoCollection.Count(new BsonDocument());

            QueryDocument query = new QueryDocument();
            BsonDocument b = new BsonDocument();
            b.Add("$ne", true);
            query.Add("isDeleted", b);
            query.Add("yearId", "591421b172a6d61d0aee0f3d");

            return mongoCollection.Find(query).ToList();
        }

        public List<BsonDocument> getAllSources(string classId)
        {
            IMongoCollection<BsonDocument> mongoCollection = mongoDataBase.GetCollection<BsonDocument>("adminEnrollTrains");//选择集合，相当于表

            QueryDocument query = new QueryDocument();
            query.Add("isSucceed", 1);
            query.Add("trainId", new ObjectId(classId));

            return mongoCollection.Find(query).ToList();
        }

        public BsonDocument getTeacher(string teacherId)
        {
            IMongoCollection<BsonDocument> mongoCollection = mongoDataBase.GetCollection<BsonDocument>("teachers");//选择集合，相当于表

            QueryDocument query = new QueryDocument();
            BsonDocument b = new BsonDocument();
            b.Add("$ne", true);
            query.Add("isDeleted", b);
            query.Add("_id", new ObjectId(teacherId));

            return mongoCollection.Find(query).FirstOrDefault();
        }

        public BsonDocument getStudent(string studentId)
        {
            IMongoCollection<BsonDocument> mongoCollection = mongoDataBase.GetCollection<BsonDocument>("studentInfos");//选择集合，相当于表

            QueryDocument query = new QueryDocument();
            BsonDocument b = new BsonDocument();
            b.Add("$ne", true);
            query.Add("isDeleted", b);
            query.Add("_id", new ObjectId(studentId));

            return mongoCollection.Find(query).FirstOrDefault();
        }

        public void Dispose() {
            
        }
    }
}
