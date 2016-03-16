using MailCRM.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MailCRM.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            //get all the comments and return to view
            List<Comment> comments = new List<Comment>();
            using (MailCRMEntities entities = new MailCRMEntities())
            {
                comments = entities.Comments.OrderByDescending(i => i.PostDate).ToList();
            }
            return View(comments);
        }

        public ActionResult Agave()
        {
            return View();
        }
    }
}