using EasyShoppingApp.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EasyShoppingApp.Controllers
{
    public class RegisterController : Controller
    {
      
        //tambahan untuk login authentication
        SqlConnection con = new SqlConnection();
        SqlCommand com = new SqlCommand();
        SqlDataReader dr;
        //akhir authentication
        Easy_ShoppingEntities db = new Easy_ShoppingEntities();
        // GET: Register
        public ActionResult InsertData()
        {
            return View();
        }
        [HttpPost]
        public ActionResult InsertData(LoginPanel model)
        {

            LoginPanel tbl = new LoginPanel();
            tbl.Username = model.Username;
            tbl.Password = model.Password;
            db.LoginPanel.Add(tbl);
            db.SaveChanges();
            ViewBag.Message ="Haii..."+ model.Username + " Akun anda sudah aktif. "  + " Pilih Login untuk masuk ";
            return View();

        }
        //tambahan untuk login sql

        void connectionString()
        {
            con.ConnectionString = @"Data Source = FANTA\SQLEXPRESS; Initial Catalog =Easy_Shopping; Integrated Security= True";
        }
        //akhir tambahan
        public ActionResult SetDataInDataBase()
        {
            return View();
        }
        [HttpPost]
        public ActionResult SetDataInDataBase(LoginPanel model)
        {
            connectionString();
            con.Open();
            com.Connection = con;
            com.CommandText = "select * from LoginPanel where username='" + model.Username + "'and password='" + model.Password + "'";
            dr = com.ExecuteReader();
            if (dr.Read())
            {
               
                con.Close();
                //tambahan untuk session login
                Session["anyName"] = dr;
                string variable = "";
                if (Session["anyName"] != null)
                    variable = Session["anyName"].ToString();
                if (Session["anyName"] != null)
                {
                    return View("Login_Success");
                }

                else
                    return View();
                    //akhir tambahan session login




                }
            else
            {
                con.Close();
                ViewBag.Message ="Username dan Password tidak terdaftar, Silahkan daftar dahulu!!";
                return View();

            }
        }
       
        public ActionResult ShowDataBaseForUser()
        {
            var item = db.LoginPanel.ToList();
            return View(item);
        }
        public ActionResult Delete(int id)
        {
            var item = db.LoginPanel.Where(x => x.ID == id).First();
            db.LoginPanel.Remove(item);
            db.SaveChanges();
            var item2 = db.LoginPanel.ToList();
            return View("ShowDataBaseForUser",item2);

           
        }
        public ActionResult Edit(int id)
        {
            var item = db.LoginPanel.Where(x => x.ID == id).First();
            return View(item);
        }

        [HttpPost]
        public ActionResult Edit(LoginPanel model)
        {
            var item = db.LoginPanel.Where(x => x.ID == model.ID).First();
            item.Username = model.Username;
            item.Password = model.Password;
            db.SaveChanges();
            //untuk mengecek setelah di update
            var item2 = db.LoginPanel.ToList();
            return View("ShowDataBaseForUser", item2);

        }

        //Untuk Export Excel
        public void ExportToExcel()
        {
            var datauser = db.LoginPanel.ToList();
           
    
            ExcelPackage pck = new ExcelPackage();

            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("DataAkun");

            ws.Cells["A1"].Value = "Username";
            ws.Cells["B1"].Value = "Password";  
         
            int rowStart = 2;
            foreach(var item in datauser)
            {
                ws.Cells[string.Format("A{0}", rowStart)].Value = item.Username;
                ws.Cells[string.Format("B{0}", rowStart)].Value = item.Password;
                rowStart++;


            }
            ws.Cells["A:AZ"].AutoFitColumns();
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment:filename=" + "ExcelReport.xlsx");
            Response.BinaryWrite(pck.GetAsByteArray());
            Response.End();



        }
    }
}