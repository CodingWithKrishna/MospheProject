using Assignment.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Ajax.Utilities;
using System.Diagnostics.SymbolStore;
using System.Web.Security;

namespace Assignment.Controllers
{
    public class LoginController : Controller
    {
        // GET: Login
        public ActionResult LogIn()
        {
            return View();
        }

        [HttpPost]
        public ActionResult LogIn(MemberLogin model)
        {
            using (var context = new MospheEntities())
            {
                bool isValid = context.registeruser.Any(x => x.Username == model.Username && x.Password == model.Password);

                if (isValid)
                {
                    FormsAuthentication.SetAuthCookie(model.Username, false);
                    ModelState.Clear();
                    return RedirectToAction("Index", "RegisterUser");
                }

                ModelState.AddModelError("", "Incorrect Username or password ");
            }           
            return View();
        }

        public  ActionResult Logout()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("LogIn");
        }

    }
}