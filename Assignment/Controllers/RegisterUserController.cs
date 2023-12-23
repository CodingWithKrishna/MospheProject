using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using Assignment;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;

namespace Assignment.Controllers
{
    [Authorize]
    public class RegisterUserController : Controller
    {
        private MospheEntities db = new MospheEntities();

        // GET: RegisterUser
        public ActionResult Index()
        {
            return View(db.registeruser.ToList());
        }

        // GET: RegisterUser/Details/5
        [Authorize(Roles ="Admin")]
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
        public ActionResult Create()
        {
            return View();
        }

        // POST: RegisterUser/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
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
        public ActionResult Import(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                ModelState.AddModelError("file", "Please select a file.");
                return View();
            }

            try
            {
                using (var stream = new MemoryStream())
                {
                    file.CopyTo(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var rowCount = worksheet.Dimension.Rows;

                        var usersToAdd = new List<registeruser>();

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var newUser = new registeruser
                            {
                                Username = worksheet.Cells[row, 1].Value?.ToString(),
                                Password = worksheet.Cells[row, 2].Value?.ToString(),
                                Email = worksheet.Cells[row, 3].Value?.ToString(),
                            };

                            usersToAdd.Add(newUser);
                        }

                        db.registeruser.AddRange(usersToAdd);
                        db.SaveChanges();
                    }
                }

                return RedirectToAction("Index", "RegisterUser"); // Redirect to a relevant page after successful import
            }
            catch (Exception ex)
            {
                ModelState.AddModelError("file", $"Error: {ex.Message}");
                return View();
            }
        }

        [HttpPost]
        public ActionResult Export()
        {
            var users = db.registeruser.ToList();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Users");
                worksheet.Cells.LoadFromCollection(users, true);

                var memoryStream = new MemoryStream();
                package.SaveAs(memoryStream);

                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Users.xlsx");
            }
        }
        [HttpPost]
        public FileResult ExportToExcel()
        {
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[3] { new DataColumn("Username"),
                                                     new DataColumn("Password"),
                                                     new DataColumn("Email"),});

            var insuranceCertificate = db.registeruser.ToList() ;

            foreach (var insurance in insuranceCertificate)
            {
                dt.Rows.Add(insurance.Username,insurance.Password,insurance.Email );
            }

            using (XLWorkbook wb = new XLWorkbook()) //Install ClosedXml from Nuget for XLWorkbook  
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream()) //using System.IO;  
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelFile.xlsx");
                }
            }
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
