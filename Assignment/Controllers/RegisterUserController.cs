using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Claims;
using System.Web;
using System.Web.Mvc;
using Assignment;
using ClosedXML.Excel;
using CsvHelper.Configuration;
using CsvHelper;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using Assignment.Models;
using System.Configuration;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Web.Configuration;
using System.CodeDom;

namespace Assignment.Controllers
{
    [Authorize]
    public class RegisterUserController : Controller
    {
        private MospheEntities db = new MospheEntities();

        // GET: RegisterUser
        public ActionResult Index()
        {
            if (User.IsInRole("Admin"))
            {

                return View(db.registeruser.ToList());
            }
            else if (User.Identity.IsAuthenticated)
            {
                var claimsIdentity = User.Identity as ClaimsIdentity;

                // Fetch user-specific data for authenticated users
                var userIdClaim = claimsIdentity.Name;

                if (User.IsInRole("User"))
                {
                    var userId = userIdClaim;
                    var userData = db.registeruser.Where(u => u.Username == userId).ToList();
                    return View(userData);
                }
            }
            else
            {
                // Handle the case when a user is not authenticated (e.g., redirect to login)
                return RedirectToAction("Login", "Account");
            }
            return View();

        }

        // GET: RegisterUser/Details/5

        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            registeruser registeruser = db.registeruser.Find(id);
            if (registeruser == null)
            {
                return HttpNotFound();
            }
            return View(registeruser);
        }

        // GET: RegisterUser/Create
        [AllowAnonymous]

        public ActionResult Create()
        {
            return View();
        }

        // POST: RegisterUser/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,Username,Password,Email")] registeruser registeruser)
        {
            if (ModelState.IsValid)
            {
                db.registeruser.Add(registeruser);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(registeruser);
        }

        // GET: RegisterUser/Edit/5
        [Authorize(Roles = "Admin")]

        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            registeruser registeruser = db.registeruser.Find(id);
            if (registeruser == null)
            {
                return HttpNotFound();
            }
            return View(registeruser);
        }

        // POST: RegisterUser/Edit/5

        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin")]

        public ActionResult Edit([Bind(Include = "Id,Username,Password,Email")] registeruser registeruser)
        {
            if (ModelState.IsValid)
            {
                db.Entry(registeruser).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(registeruser);
        }

        // GET: RegisterUser/Delete/5
        [Authorize(Roles = "Admin")]

        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            registeruser registeruser = db.registeruser.Find(id);
            if (registeruser == null)
            {
                return HttpNotFound();
            }
            return View(registeruser);
        }

        // POST: RegisterUser/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            registeruser registeruser = db.registeruser.Find(id);
            db.registeruser.Remove(registeruser);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        [HttpGet]
        public ActionResult Import()
        {
            return View();
        }

        [HttpPost]
      
        public FileResult ExportToExcel()
        {
            try {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                DataTable dt = new DataTable("Grid");
                dt.Columns.AddRange(new DataColumn[3]
                {
        new DataColumn("Username"),
        new DataColumn("Password"),
        new DataColumn("Email"),
                });
                List<registeruser> insuranceCertificate;
                if (User.IsInRole("User") && !User.IsInRole("Admin"))
                {
                    string username = GetUserLoginUsername();
                    insuranceCertificate = db.registeruser.Where(x => x.Username == username).ToList();
                }
                else
                {
                    insuranceCertificate = db.registeruser.ToList();
                }
                foreach (var insurance in insuranceCertificate)
                {
                    dt.Rows.Add(insurance.Username, insurance.Password, insurance.Email);
                }

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Grid");
                    worksheet.Cells.LoadFromDataTable(dt, true);

                    using (var stream = new MemoryStream())
                    {
                        package.SaveAs(stream);
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelFile.xlsx");
                    }
                }
            }
            catch(Exception ex) { throw; }
            
        }
        [HttpPost]
        public ActionResult Import(HttpPostedFileBase file)
        {
            try
            {
                if (file != null && file.ContentLength > 0)
                {
                    ExcelPackage.LicenseContext = LicenseContext.Commercial;

                    using (var package = new ExcelPackage(file.InputStream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var username = worksheet.Cells[row, 1].Text;
                            var password = worksheet.Cells[row, 2].Text;
                            var email = worksheet.Cells[row, 3].Text;

                            var newUser = new registeruser
                            {
                                Username = username,
                                Password = password,
                                Email = email
                                // Additional properties...
                            };

                            db.registeruser.Add(newUser);
                        }

                        db.SaveChanges();
                    }
                }

                return RedirectToAction("Index");
            }
            catch(Exception ex) { throw; }
        }
        private string GetUserLoginUsername()
        {
            var claimsIdentity = User.Identity as ClaimsIdentity;

            // Fetch user-specific data for authenticated users
            var userIdClaim = claimsIdentity.Name;
            return userIdClaim;
        }
    
    protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    } 
}
