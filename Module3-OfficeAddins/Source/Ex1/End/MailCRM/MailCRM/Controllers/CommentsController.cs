using MailCRM.Models;
using Microsoft.AspNet.SignalR;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace MailCRM.Controllers
{
    public class CommentsController : ApiController
    {
        [Route("api/Comments")]
        public IEnumerable<Comment> Get(string email)
        {
            List<Comment> comments = new List<Comment>();
            using (MailCRMEntities entities = new MailCRMEntities())
            {
                email = email.ToLower();
                comments = entities.Comments.Where(i => i.ContactEmail == email)
                    .OrderByDescending(i => i.PostDate)
                    .ToList();
            }

            return comments;
        }

        [Route("api/Comments")]
        [HttpPost]
        public void Post(Comment comment)
        {
            using (MailCRMEntities entities = new MailCRMEntities())
            {
                comment.PostDate = DateTime.Now;
                entities.Comments.Add(comment);
                entities.SaveChanges();

                //notify the client through the hub
                var hubContext = GlobalHost.ConnectionManager.GetHubContext<CommentsHub>();
                hubContext.Clients.All.commentReceived(comment);
            }
        }
    }
}
