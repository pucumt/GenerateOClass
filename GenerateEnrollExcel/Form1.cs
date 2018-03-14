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
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Data.OleDb;
using System.Threading;
using System.Data.SqlClient;

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
            var blockLength = 7;
            var toBlock = 0;

            for (var j =0; j<courses.Rows.Count; j++)
            {
                var curClass = courses.Rows[j];
                
                if (string.IsNullOrEmpty(curClass["teacherId"].ToString()))
                {
                    continue;
                }

                //小语文
                //var xiaoyu = new List<string>();
                //xiaoyu.Add("小学一年级");
                //xiaoyu.Add("小学二年级");
                //xiaoyu.Add("小学三年级");
                //xiaoyu.Add("小学四年级");
                //xiaoyu.Add("小学五年级");
                //xiaoyu.Add("小学六年级");
                //if (curClass.GetValue("subjectName").ToString() != "语文" || (!xiaoyu.Contains(curClass.GetValue("gradeName").ToString())))
                //{
                //    continue;
                //}

                //中英语
                //var zhongying = new List<string>();
                //zhongying.Add("初中一年级");
                //zhongying.Add("初中二年级");
                //zhongying.Add("初中三年级");
                //if (curClass.GetValue("subjectName").ToString() != "英语" || (!zhongying.Contains(curClass.GetValue("gradeName").ToString())))
                //{
                //    continue;
                //}

                XSSFWorkbook hssfworkbook = null;
                using (FileStream fs = File.Open(@"template.xlsx", FileMode.Open,
                FileAccess.Read, FileShare.ReadWrite))
                {
                    //把xls文件读入workbook变量里，之后就可以关闭了  
                    hssfworkbook = new XSSFWorkbook(fs);
                    fs.Close();
                }

                var fileName = string.Format("沟通_{0}_{1}_{2}_{3}.xlsx", curClass["schoolArea"].ToString(), curClass["subjectName"].ToString(), curClass["gradeName"].ToString(), curClass["name"].ToString());

                XSSFSheet sheet1 = hssfworkbook.GetSheet("template") as XSSFSheet;
                hssfworkbook.SetSheetName(0, curClass["name"].ToString());

                var curSource = dbHelper.getAllSources(curClass["_id"].ToString());
                var curTeacher = dbHelper.getTeacher(curClass["teacherId"].ToString()).Rows[0];

                for (var i = 0; i < curSource.Rows.Count; i++)
                {
                    var curStudent = dbHelper.getStudent(curSource.Rows[i]["studentId"].ToString()).Rows[0];

                    sheet1.GetRow(2 + blockLength * i).GetCell(1).SetCellValue(curClass["name"].ToString());
                    sheet1.GetRow(2 + blockLength * i).GetCell(2).SetCellValue(curStudent["name"].ToString());
                    sheet1.GetRow(2 + blockLength * i).GetCell(3).SetCellValue(curClass["teacherName"].ToString());
                    sheet1.GetRow(2 + blockLength * i).GetCell(4).SetCellValue(curStudent["mobile"].ToString());
                    sheet1.GetRow(2 + blockLength * i).GetCell(5).SetCellValue(curTeacher["mobile"].ToString());
                }

                sheet1.ForceFormulaRecalculation = true;

                using (FileStream fileStream = File.Open(fileName,
                    FileMode.Create, FileAccess.ReadWrite))
                {
                    hssfworkbook.Write(fileStream);
                    fileStream.Close();
                }
                //if (toBlock == 3)
                //{
                //    break;
                //}

                //toBlock++;
            }
            MessageBox.Show("导出成功！");
        }

        //public List<Order> ExcelToDS(string Path)
        //{
        //    FileStream file = new FileStream(Path, FileMode.Open, FileAccess.Read);
        //    XSSFWorkbook hssfworkbook = new XSSFWorkbook(file);
        //    XSSFSheet sheet1 = hssfworkbook.GetSheet("报名情况") as XSSFSheet;
        //    List<Order> orders = new List<Order>();
        //    var i = 1;
        //    while(sheet1.GetRow(i)!=null)
        //    {
        //        var newOrder = new Order();
        //        for (var j = 0; j < 11; j++)
        //        {
        //            newOrder.studentName = sheet1.GetRow(i).GetCell(0).StringCellValue;
        //            newOrder.mobile = sheet1.GetRow(i).GetCell(1).StringCellValue;
        //            newOrder.className = sheet1.GetRow(i).GetCell(8).StringCellValue;
        //            newOrder.school = sheet1.GetRow(i).GetCell(7).StringCellValue;
        //            newOrder.grade = sheet1.GetRow(i).GetCell(10).StringCellValue;
        //            newOrder.subject = sheet1.GetRow(i).GetCell(6).StringCellValue;
        //            newOrder.courseTime = sheet1.GetRow(i).GetCell(9).StringCellValue;
        //            newOrder.studentSchool = (sheet1.GetRow(i).GetCell(2)!=null?sheet1.GetRow(i).GetCell(2).StringCellValue:"");
        //            newOrder.studentClass = (sheet1.GetRow(i).GetCell(3)!=null?sheet1.GetRow(i).GetCell(3).StringCellValue:"");
        //        }
        //        if(newOrder.className!="")
        //        {
        //            orders.Add(newOrder);
        //        }                
        //        i++;
        //    }

        //     return orders;
        //}
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
        string connStr = "server=127.0.0.1;Database=website2;UID=root;PWD=root";
        //string database = "websiteOnline20";
        //MongoClient mongodb;
        //IMongoDatabase mongoDataBase;

        
        public DBHelper() {
            //mongodb = new MongoClient(conn);//连接数据库
            //mongoDataBase = mongodb.GetDatabase(database);//选择数据库名
        }

        public DataTable getAllCourses()
        {
            //// 小语文 中英语 特殊
            //IMongoCollection<BsonDocument> mongoCollection = mongoDataBase.GetCollection <BsonDocument> ("trainClasss");//选择集合，相当于表

            ////var count = mongoCollection.Count(new BsonDocument());

            //QueryDocument query = new QueryDocument();
            //BsonDocument b = new BsonDocument();
            //b.Add("$ne", true);
            //query.Add("isDeleted", b);
            //query.Add("yearId", "597586d3148d30515a12bf06");

            //return mongoCollection.Find(query).ToList();

            // 春季课程    
            string sql = "select * from trainClasss where isDeleted=false and yearId='2d84ae1f5b5293e811e7dbdb298e4580' and attributeId='2d84ae1f5b5293e811e7dbdb482da080' ";
            MySqlDataAdapter myda = new MySqlDataAdapter(sql, this.connStr);
            DataSet ds = new DataSet();
            myda.Fill(ds);

            myda.Dispose();
            return ds.Tables[0];
        }

        public DataTable getAllSources(string classId)
        {
            //IMongoCollection<BsonDocument> mongoCollection = mongoDataBase.GetCollection<BsonDocument>("adminEnrollTrains");//选择集合，相当于表

            //QueryDocument query = new QueryDocument();
            //query.Add("isSucceed", 1);
            //query.Add("trainId", new ObjectId(classId));

            //return mongoCollection.Find(query).ToList();

            string sql = "select * from adminEnrollTrains where isDeleted=false and isSucceed=1 and trainId='"+ classId + "' ";
            MySqlDataAdapter myda = new MySqlDataAdapter(sql, this.connStr);
            DataSet ds = new DataSet();
            myda.Fill(ds);

            myda.Dispose();
            return ds.Tables[0];
        }

        public DataTable getTeacher(string teacherId)
        {
            //IMongoCollection<BsonDocument> mongoCollection = mongoDataBase.GetCollection<BsonDocument>("teachers");//选择集合，相当于表

            //QueryDocument query = new QueryDocument();
            //BsonDocument b = new BsonDocument();
            //b.Add("$ne", true);
            //query.Add("isDeleted", b);
            //query.Add("_id", new ObjectId(teacherId));

            //return mongoCollection.Find(query).FirstOrDefault();

            string sql = "select * from teachers where isDeleted=false ";
            MySqlDataAdapter myda = new MySqlDataAdapter(sql, this.connStr);
            DataSet ds = new DataSet();
            myda.Fill(ds);

            myda.Dispose();
            return ds.Tables[0];
        }

        public DataTable getStudent(string studentId)
        {
            //IMongoCollection<BsonDocument> mongoCollection = mongoDataBase.GetCollection<BsonDocument>("studentInfos");//选择集合，相当于表

            //QueryDocument query = new QueryDocument();
            //BsonDocument b = new BsonDocument();
            //b.Add("$ne", true);
            //query.Add("isDeleted", b);
            //query.Add("_id", new ObjectId(studentId));

            //return mongoCollection.Find(query).FirstOrDefault();
            string sql = "select * from studentInfos where isDeleted=false and _id='"+ studentId + "'";
            MySqlDataAdapter myda = new MySqlDataAdapter(sql, this.connStr);
            DataSet ds = new DataSet();
            myda.Fill(ds);

            myda.Dispose();
            return ds.Tables[0];
        }

        public void Dispose() {
            
        }
    }
}
