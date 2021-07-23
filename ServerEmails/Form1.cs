using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Exchange.WebServices.Data;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.IO;

namespace ServerEmails
{
    

    public partial class Form1 : Form
    {

        SqlConnection conRegister = new SqlConnection("Data Source=10.240.107.205, 1433;Initial Catalog=Registers;User ID=sa;Password=CelineDion@50;");
        SqlConnection conLeave = new SqlConnection("Data Source=10.240.107.205, 1433;Initial Catalog=LeaveApplications;User ID=sa;Password=CelineDion@50;");
        SqlConnection conUsers = new SqlConnection("Data Source=10.240.107.205,1433;Initial Catalog=NCDOEChatDB;User ID=sa;Password=CelineDion@50");
        SqlConnection conAudit = new SqlConnection("Data Source=10.240.107.205, 1433;Initial Catalog=Audit;User ID=sa;Password=CelineDion@50;");
        SqlCommand cmd;
        SqlDataAdapter adapt;



        DataTable dt = new DataTable();


        string email = "";

        int reference;
        string subject;
        string body;
        int applicant;


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer2.Start();
            timer1.Start();
            label1.Text = "Started!";
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            timer2.Stop();
            timer1.Stop();
            label1.Text = "Stoped!";
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            timer2.Start();
            timer1.Start();
            label1.Text = "Started!";
   
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            DateTime dt10AM = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 10, 0, 0);
            DateTime dt4PM = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 16, 0, 0);

            if (DateTime.Now > dt10AM && DateTime.Now < dt4PM)
            {

                string component = "";
                string email = "";


                conLeave.Open();
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("select *  from Emails", conLeave);
                adapt.Fill(dt);
                conLeave.Close();

                int countRows = dt.Rows.Count;

                if (countRows > 0)
                {

                    for (int i = 0; i < countRows; i++)
                    {
                        reference = (int)dt.Rows[i]["Reference"];
                        subject = dt.Rows[i]["Subject"].ToString();
                        body = dt.Rows[i]["Body"].ToString();
                        applicant = (int)dt.Rows[i]["Applicant"];

                        conUsers.Open();
                        DataTable dtUsers = new DataTable();
                        adapt = new SqlDataAdapter("select Component, Supervisor from Users where Persal = " + applicant, conUsers);
                        adapt.Fill(dtUsers);
                        conUsers.Close();

                        component = dtUsers.Rows[0]["Component"].ToString();

                        int supervisor;

                        if (dtUsers.Rows[0]["Supervisor"] != DBNull.Value)
                        {
                            supervisor = (int)dtUsers.Rows[0]["Supervisor"];



                        }
                        else
                        {
                            supervisor = 0;
                        }




                        if (supervisor != 0)
                        {
                            conUsers.Open();
                            DataTable dtEmail = new DataTable();
                            adapt = new SqlDataAdapter("select Email from Users where Persal = " + supervisor, conUsers);
                            adapt.Fill(dtEmail);
                            conUsers.Close();

                            email = dtEmail.Rows[0]["Email"].ToString();


                        }
                        else
                        {
                            conUsers.Open();
                            DataTable dtOtherEmail = new DataTable();
                            adapt = new SqlDataAdapter("select Email  from Users where AppRole = 'ADMIN2' AND Component = '" + component + "' ", conUsers);
                            adapt.Fill(dtOtherEmail);
                            conUsers.Close();
                            email = dtOtherEmail.Rows[0]["Email"].ToString();
                        }

                        //SEND EMAIL
                        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                        //service.AutodiscoverUrl("omphemetsengakaemang@ncdoe.gov.za");

                        service.Url = new Uri("https://mail.ncdoe.gov.za/ews/Exchange.asmx");

                        service.UseDefaultCredentials = true;
                        service.Credentials = new WebCredentials("ncdoesystemsnotifica", "12345678@D");


                        EmailMessage message = new EmailMessage(service);
                        message.Subject = subject;
                        message.Body = body;

                        message.ToRecipients.Add(email);
                        message.Save();

                        message.SendAndSaveCopy();




                        using (SqlConnection conLeave = new SqlConnection("Data Source=10.240.107.205, 1433;Initial Catalog=LeaveApplications;User ID=sa;Password=CelineDion@50;"))
                        {
                            conLeave.Open();
                            using (SqlCommand command = new SqlCommand("DELETE FROM EMAILS WHERE Reference = " + reference, conLeave))
                            {
                                command.ExecuteNonQuery();
                            }
                            conLeave.Close();
                        }


                    }

                }

                //////////////////////////////////////////////////////////////////////////////////////
                ///

                conLeave.Open();
                DataTable dtRecommended = new DataTable();
                adapt = new SqlDataAdapter("select *  from RecommendedEmails", conLeave);
                adapt.Fill(dtRecommended);
                conLeave.Close();

                int countRowsRecommended = dtRecommended.Rows.Count;


                string recommendedComponent = "";

                if (countRowsRecommended > 0)
                {

                    for (int i = 0; i < countRowsRecommended; i++)
                    {
                        reference = (int)dtRecommended.Rows[i]["Reference"];
                        subject = dtRecommended.Rows[i]["Subject"].ToString();
                        body = dtRecommended.Rows[i]["Body"].ToString();
                        applicant = (int)dtRecommended.Rows[i]["Applicant"];
                        recommendedComponent = dtRecommended.Rows[i]["Component"].ToString();

                        if(recommendedComponent == "EMIS")
                        {
                            recommendedComponent = "ICT";
                        }
                        conUsers.Open();
                        DataTable dtOtherEmail = new DataTable();
                        adapt = new SqlDataAdapter("select Email  from Users where AppRole = 'ADMIN1' AND Component = '" + recommendedComponent + "' ", conUsers);
                        adapt.Fill(dtOtherEmail);
                        conUsers.Close();
                        email = dtOtherEmail.Rows[0]["Email"].ToString();




                        //SEND EMAIL
                        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                        //service.AutodiscoverUrl("omphemetsengakaemang@ncdoe.gov.za");

                        service.Url = new Uri("https://mail.ncdoe.gov.za/ews/Exchange.asmx");

                        service.UseDefaultCredentials = true;
                        service.Credentials = new WebCredentials("ncdoesystemsnotifica", "12345678@D");


                        EmailMessage message = new EmailMessage(service);
                        message.Subject = subject;
                        message.Body = body;

                        message.ToRecipients.Add(email);
                        message.Save();

                        message.SendAndSaveCopy();




                        using (SqlConnection conLeave = new SqlConnection("Data Source=10.240.107.205, 1433;Initial Catalog=LeaveApplications;User ID=sa;Password=CelineDion@50;"))
                        {
                            conLeave.Open();
                            using (SqlCommand command = new SqlCommand("DELETE FROM RecommendedEmails WHERE Reference = " + reference, conLeave))
                            {
                                command.ExecuteNonQuery();
                            }
                            conLeave.Close();
                        }
                    }
                }
                ////////////////SEND NEW QUESTION EMAIL//////////////////////////////////////////////////////////////////////////////////////////////////////
                conAudit.Open();
                DataTable dtAuditNew = new DataTable();
                adapt = new SqlDataAdapter("select *  from Questions where Email = 'SEND' AND Progress = 'NEW'", conAudit);
                adapt.Fill(dtAuditNew);
                conAudit.Close();

                List<int> responsibleOfficial = new List<int>();

                foreach (DataRow row in dtAuditNew.Rows)
                {
                    

                    if(responsibleOfficial.Contains(Convert.ToInt32(row[3])))
                    {

                    }
                    else
                    {
                        responsibleOfficial.Add(Convert.ToInt32(row[3]));
                    }
                }

                foreach(var item in responsibleOfficial)
                {
                    conUsers.Open();
                    DataTable dtUsers = new DataTable();
                    adapt = new SqlDataAdapter("select Email from Users where Persal =" + item, conUsers);
                    adapt.Fill(dtUsers);
                    conUsers.Close();




                    //SEND EMAIL
                    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                    //service.AutodiscoverUrl("omphemetsengakaemang@ncdoe.gov.za");

                    service.Url = new Uri("https://mail.ncdoe.gov.za/ews/Exchange.asmx");

                    service.UseDefaultCredentials = true;
                    service.Credentials = new WebCredentials("ncdoesystemsnotifica", "12345678@D");


                    EmailMessage message = new EmailMessage(service);
                    message.Subject = "NEW AUDIT QUESTION";
                    message.Body = "There are new audit questions for you to answer, please log into the office management system to attend to them.";

                    message.ToRecipients.Add(dtUsers.Rows[0]["Email"].ToString());
                    message.Save();

                    message.SendAndSaveCopy();
                }

                using (SqlConnection conAudit = new SqlConnection("Data Source=10.240.107.205, 1433;Initial Catalog=Audit;User ID=sa;Password=CelineDion@50;"))
                {
                    conAudit.Open();
                    using (SqlCommand command = new SqlCommand("UPDATE Questions set Email = 'NO NEED'  WHERE Progress = 'NEW' AND Email = 'SEND'  " , conAudit))
                    {
                        command.ExecuteNonQuery();
                    }
                    conAudit.Close();
                }


                ///////////////SEND SENT BACK EMAIL/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                conAudit.Open();
                DataTable dtAuditSentBack = new DataTable();
                adapt = new SqlDataAdapter("select *  from Questions where Email = 'SEND' AND Progress = 'SENT BACK'", conAudit);
                adapt.Fill(dtAuditSentBack);
                conAudit.Close();

                responsibleOfficial = new List<int>();

                foreach (DataRow row in dtAuditSentBack.Rows)
                {


                    if (responsibleOfficial.Contains(Convert.ToInt32(row[3])))
                    {

                    }
                    else
                    {
                        responsibleOfficial.Add(Convert.ToInt32(row[3]));
                    }
                }

                foreach (var item in responsibleOfficial)
                {
                    conUsers.Open();
                    DataTable dtUsers = new DataTable();
                    adapt = new SqlDataAdapter("select Email from Users where Persal =" + item, conUsers);
                    adapt.Fill(dtUsers);
                    conUsers.Close();




                    //SEND EMAIL
                    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                    //service.AutodiscoverUrl("omphemetsengakaemang@ncdoe.gov.za");

                    service.Url = new Uri("https://mail.ncdoe.gov.za/ews/Exchange.asmx");

                    service.UseDefaultCredentials = true;
                    service.Credentials = new WebCredentials("ncdoesystemsnotifica", "12345678@D");


                    EmailMessage message = new EmailMessage(service);
                    message.Subject = "AUDIT QUESTION SENT BACK";
                    message.Body = "One or more of your answered audit questions have been sent back, please log into the office management system to attend to them.";

                    message.ToRecipients.Add(dtUsers.Rows[0]["Email"].ToString());
                    message.Save();

                    message.SendAndSaveCopy();
                }

                using (SqlConnection conAudit = new SqlConnection("Data Source=10.240.107.205, 1433;Initial Catalog=Audit;User ID=sa;Password=CelineDion@50;"))
                {
                    conAudit.Open();
                    using (SqlCommand command = new SqlCommand("UPDATE Questions set Email = 'NO NEED'  WHERE Progress = 'SENT BACK' AND Email = 'SEND'  " , conAudit))
                    {
                        command.ExecuteNonQuery();
                    }
                    conAudit.Close();
                }
                ///////////////SEND ANSWERED EMAIL/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                conAudit.Open();
                DataTable dtAuditAnswered = new DataTable();
                adapt = new SqlDataAdapter("select *  from Questions where Email = 'SEND' AND Progress = 'ANSWERED'", conAudit);
                adapt.Fill(dtAuditSentBack);
                conAudit.Close();

                responsibleOfficial = new List<int>();

                foreach (DataRow row in dtAuditSentBack.Rows)
                {


                    if (responsibleOfficial.Contains(Convert.ToInt32(row[6])))
                    {

                    }
                    else
                    {
                        responsibleOfficial.Add(Convert.ToInt32(row[6]));
                    }
                }

                foreach (var item in responsibleOfficial)
                {
                    conUsers.Open();
                    DataTable dtUsers = new DataTable();
                    adapt = new SqlDataAdapter("select Email from Users where Persal =" + item, conUsers);
                    adapt.Fill(dtUsers);
                    conUsers.Close();




                    //SEND EMAIL
                    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                    //service.AutodiscoverUrl("omphemetsengakaemang@ncdoe.gov.za");

                    service.Url = new Uri("https://mail.ncdoe.gov.za/ews/Exchange.asmx");

                    service.UseDefaultCredentials = true;
                    service.Credentials = new WebCredentials("ncdoesystemsnotifica", "12345678@D");


                    EmailMessage message = new EmailMessage(service);
                    message.Subject = "ANSWERED AUDIT QUESTION";
                    message.Body = "One or more of your audit questions have been answered, please log into the office management system to attend to them.";

                    message.ToRecipients.Add(dtUsers.Rows[0]["Email"].ToString());
                    message.Save();

                    message.SendAndSaveCopy();
                }

                using (SqlConnection conAudit = new SqlConnection("Data Source=10.240.107.205, 1433;Initial Catalog=Audit;User ID=sa;Password=CelineDion@50;"))
                {
                    conAudit.Open();
                    using (SqlCommand command = new SqlCommand("UPDATE Questions set Email = 'NO NEED'  WHERE Progress = 'ANSWERED' AND Email = 'SEND'  ", conAudit))
                    {
                        command.ExecuteNonQuery();
                    }
                    conAudit.Close();
                }




            }



        }

        private void label1_Click(object sender, EventArgs e)
        {


            

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            DateTime dt350AM = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 15,51, 0);
            DateTime dt351PM = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 15, 52, 0);

            DataTable dw = new DataTable();
            DataTable dc = new DataTable();
            DataTable dp = new DataTable();

            SqlDataAdapter da = new SqlDataAdapter();

            if (DateTime.Now >= dt350AM && DateTime.Now <= dt351PM)
            {
                dw = new DataTable();
                da = new SqlDataAdapter("select NCDOEChatDB.dbo.Users.UserName as Name, WorkingDays.Persal, WorkingDays.Start, WorkingDays.[end] , WorkingDays.NumberOfDays, WorkingDays.[Type], NCDOEChatDB.dbo.Users.Component From LeaveApplications.dbo.WorkingDays JOIN NCDOEChatDB.dbo.Users ON LeaveApplications.dbo.WorkingDays.Persal = NCDOEChatDB.dbo.Users.Persal where ApproveDate = (cast (GETDATE() AS Date))", conLeave);
                da.Fill(dw);
                conLeave.Close();



                dc = new DataTable();
                da = new SqlDataAdapter("select NCDOEChatDB.dbo.Users.UserName as Name, CalendarDays.Persal, CalendarDays.Start, CalendarDays.[end] , CalendarDays.NumberOfDays, CalendarDays.NumberOfMonths, CalendarDays.[Type], NCDOEChatDB.dbo.Users.Component From LeaveApplications.dbo.CalendarDays JOIN NCDOEChatDB.dbo.Users ON LeaveApplications.dbo.CalendarDays.Persal = NCDOEChatDB.dbo.Users.Persal where ApproveDate = (cast (GETDATE() AS Date)) ", conLeave);
                da.Fill(dc);
                conLeave.Close();



                dp = new DataTable();
                da = new SqlDataAdapter("select NCDOEChatDB.dbo.Users.UserName as Name, PartsOfDay.Persal, PartsOfDay.LeaveDate, PartsOfDay.StartTime, PartsOfDay.EndTime , PartsOfDay.NumberOfHours, PartsOfDay.NumberOfMinutes, PartsOfDay.[Type], NCDOEChatDB.dbo.Users.Component From LeaveApplications.dbo.PartsOfDay JOIN NCDOEChatDB.dbo.Users ON LeaveApplications.dbo.PartsOfDay.Persal = NCDOEChatDB.dbo.Users.Persal where ApproveDate = (cast (GETDATE() AS Date)) ", conLeave);
                da.Fill(dp);
                conLeave.Close();
                dt = new DataTable();
                dt.Clear();
                dt.Columns.Add("NAME");
                dt.Columns.Add("PERSAL.NO");
                dt.Columns.Add("DATE");
                dt.Columns.Add("NO. OF DAYS/HOURS");
                dt.Columns.Add("TYPE OF LEAVE");


                //----ICT and EMIS----/////////////////////////////////////////////////////////////////////////////////////////////////
                foreach (DataRow row in dw.Rows)
                {
                    if (row[6].ToString() == "ICT" || row[6].ToString() == "EMIS")
                    {
                        DataRow dataRow = dt.NewRow();

                        dataRow["NAME"] = row[0];
                        dataRow["PERSAL.NO"] = row[1];
                        dataRow["DATE"] = row[2].ToString().Substring(0, row[2].ToString().Length - 8) + " - " + row[3].ToString().Substring(0, row[3].ToString().Length - 8);
                        dataRow["NO. OF DAYS/HOURS"] = row[4].ToString() + " Days";
                        dataRow["TYPE OF LEAVE"] = row[5];

                        dt.Rows.Add(dataRow);
                    }
                }
                foreach (DataRow row in dc.Rows)
                {
                    if (row[7].ToString() == "ICT" || row[7].ToString() == "EMIS")
                    {
                        DataRow dataRow = dt.NewRow();

                        dataRow["NAME"] = row[0];
                        dataRow["PERSAL.NO"] = row[1];
                        dataRow["DATE"] = row[2].ToString().Substring(0, row[2].ToString().Length - 8) + " - " + row[3].ToString().Substring(0, row[3].ToString().Length - 8);
                        if (row[4].ToString() == "0")
                        {
                            dataRow["NO. OF DAYS/HOURS"] = row[5].ToString() + " Months";
                        }
                        else
                        {
                            dataRow["NO. OF DAYS/HOURS"] = row[4].ToString() + " Days";
                        }

                        dataRow["TYPE OF LEAVE"] = row[5];

                        dt.Rows.Add(dataRow);
                    }
                }
                foreach (DataRow row in dp.Rows)
                {
                    if (row[8].ToString() == "ICT" || row[8].ToString() == "EMIS")
                    {
                        DataRow dataRow = dt.NewRow();

                        dataRow["NAME"] = row[0];
                        dataRow["PERSAL.NO"] = row[1];
                        dataRow["DATE"] = row[2].ToString().Substring(0, row[2].ToString().Length - 8) + "[" + row[3].ToString().Substring(0, row[3].ToString().Length - 3) + " - " + row[4].ToString().Substring(0, row[4].ToString().Length - 3) + "]";
                        dataRow["NO. OF DAYS/HOURS"] = row[5].ToString() + " Hrs " + row[6].ToString() + " mins";
                        dataRow["TYPE OF LEAVE"] = row[7];

                        dt.Rows.Add(dataRow);
                    }
                }
                Print("ICTEMIS");


                conUsers.Open();
                DataTable edt = new DataTable();
                adapt = new SqlDataAdapter("select Email from Users where Component = 'ICT' AND AppRole = 'SECR'", conUsers);
                adapt.Fill(edt);
                conUsers.Close();

                email = edt.Rows[0]["Email"].ToString();




                //SEND EMAIL
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                //service.AutodiscoverUrl("omphemetsengakaemang@ncdoe.gov.za");

                service.Url = new Uri("https://mail.ncdoe.gov.za/ews/Exchange.asmx");

                service.UseDefaultCredentials = true;
                service.Credentials = new WebCredentials("ncdoesystemsnotifica", "12345678@D");


                EmailMessage message = new EmailMessage(service);
                message.Subject = "TODAY'S APPROVED LEAVES";
                message.Body = "List of People who's leaves were approved today.";
                message.Attachments.AddFileAttachment(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\LEAVES\\LeaveRegister" + "ICTEMIS" + DateTime.Now.Month + DateTime.Now.Day + ".docx");

                message.ToRecipients.Add(email);
                message.Save();

                message.SendAndSaveCopy();
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            }

        }






        private void CreateAndMoveTemplate(string unit)
        {

            if (Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\LEAVES\\"))
            {
                string sourceFile = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, @"C:\Program Files (x86)\Template\leaveRegTemplate.docx");
                string destinationFile = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\LEAVES\\LeaveRegister" + unit + DateTime.Now.Month + DateTime.Now.Day + ".docx";


                System.IO.File.Copy(sourceFile, destinationFile);
            }
            else
            {
                DirectoryInfo di = Directory.CreateDirectory(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\LEAVES\\");
                string sourceFile = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, @"C:\Program Files (x86)\Template\leaveRegTemplate.docx");
                string destinationFile = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\LEAVES\\LeaveRegister" + unit + DateTime.Now.Month + DateTime.Now.Day + ".docx";


                System.IO.File.Copy(sourceFile, destinationFile);

            }
        }





        private void Print(string unit)
        {
            try
            {
                
                CreateAndMoveTemplate(unit);
                string file = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\LEAVES\\LeaveRegister" + unit + DateTime.Now.Month + DateTime.Now.Day + ".docx";
                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                var app = new Word.Application();
                var doc = app.Documents.Open(file);

                var range = doc.Range();


                range.Find.Execute(FindText: "<unit>", Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: unit);
                range.Find.Execute(FindText: "<date>", Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: DateTime.Now.Day + " " + DateTime.Now.ToString("MMMM") + " " + DateTime.Now.Year);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    range.Find.Execute(FindText: "<name" + (i + 1).ToString() + ">", Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: dt.Rows[i]["NAME"]);
                    range.Find.Execute(FindText: "<persal" + (i + 1).ToString() + ">", Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: dt.Rows[i]["PERSAL.NO"]);
                    range.Find.Execute(FindText: "<date" + (i + 1).ToString() + ">", Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: dt.Rows[i]["DATE"]);
                    range.Find.Execute(FindText: "<days" + (i + 1).ToString() + ">", Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: dt.Rows[i]["NO. OF DAYS/HOURS"]);
                    range.Find.Execute(FindText: "<type" + (i + 1).ToString() + ">", Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: dt.Rows[i]["TYPE OF LEAVE"]);
                }


                Microsoft.Office.Interop.Word.Table table = doc.Tables[1];

                if (dt.Rows.Count < 15)
                {
                    for (int i = (dt.Rows.Count + 1); i <= 15; i++)
                    {
                        table.Rows[dt.Rows.Count + 2].Delete();
                    }
                }

                range.Find.Execute(FindText: "<noApproved1>", Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: dt.Rows.Count);
                range.Find.Execute(FindText: "<noSubmitted1>", Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: dt.Rows.Count);

                var shapes = doc.Shapes;

                foreach (Word.Shape shape in shapes)
                {
                    var initialText = shape.TextFrame.TextRange.Text;
                    var resultingText = initialText.Replace("<unit>", unit);
                    shape.TextFrame.TextRange.Text = resultingText;

                }






                doc.Save();
                //doc.SaveAs2(@"C:\Users\Ngakaemang\Documents\folder\cikti.docx");
                doc.Close();

                Marshal.ReleaseComObject(app);
                Marshal.ReleaseComObject(doc);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
