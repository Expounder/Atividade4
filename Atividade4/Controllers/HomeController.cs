using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;
using System.Net;
using System.Net.Mail;

namespace Atividade4.Controllers
{
    public class HomeController : Controller
    {
        public JsonResult Envio()
        {

            var files = Request.Files[0];
            var fileName = Path.GetFileName(files.FileName);
            var path = Path.Combine(Server.MapPath("~/Arquivos"), fileName);
            files.SaveAs(path);

            List<Entidades> users = new List<Entidades>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var package = new ExcelPackage(new FileInfo(path));
            var workbook = package.Workbook;
            var sheets = workbook.Worksheets[0];

            var excel = workbook.Worksheets.FirstOrDefault();

            int contaColuna = excel.Dimension.End.Column;
            int contaLinha = excel.Dimension.End.Row;

            for (int row = 2; row <= contaLinha; row++)
            {
                Entidades usuario = new Entidades();
                for (int col = 1; col <= contaColuna; col++)
                {
                    if (col == 1)
                    {
                        usuario.Nome = excel.Cells[row, col].Value?.ToString();
                    }
                    else if (col == 2)
                    {
                        usuario.Email = excel.Cells[row, col].Value?.ToString();
                    }
                }
                usuario.Senha = Guid.NewGuid().ToString().Replace("-", "");
                users.Add(usuario);
            }
            SendMail(users);
            return Json(new { success = true });
        }
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public void SendMail(List<Entidades> entidades)
        {
            foreach (var item in entidades)
            {
                try
                {
                    if(item.Email == null)
                    {
                        continue;
                    }
                    MailMessage mail = new MailMessage();
                    mail.To.Add(item.Email);
                    mail.From = new MailAddress("AutoMailReplace@gmail.com");
                    mail.Subject = "Recuperação de Login";
                    string Body = $"Ola {item.Nome}, o seu E-mail é: {item.Email} e a sua senha: {item.Senha} ";
                    mail.Body = Body;
                    mail.IsBodyHtml = true;

                    //Instância smtp do servidor, neste caso o gmail.
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "smtp.gmail.com";
                    smtp.Port = 587;
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = new NetworkCredential("AutoMailReplace@gmail.com", "123@utomatic");// Login e senha do e-mail.
                    smtp.EnableSsl = true;
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.Send(mail);
                }
                catch (Exception ex)
                {

                    throw;
                }

            }
        }
    }
}

public class Entidades
{
    public string Nome { get; set; }
    public string Email { get; set; }
    public string Senha { get; set; }



}
