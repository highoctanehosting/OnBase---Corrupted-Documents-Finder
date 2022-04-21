using System;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Text;

namespace DG99_Finder
{
    class Program
    {
        static readonly String hostName = Dns.GetHostName();
        static readonly String connectionString = "Server=" + hostName + ";Integrated Security=true;Database=OnBase;TrustServerCertificate=true";
        static readonly SqlConnection sqlConnection = new SqlConnection(connectionString);
        static readonly DateTime date;
        static readonly String today = date.ToString("MM/dd/yy");
        static readonly String email = "wvalerts@teamnorthwoods.com";
        static readonly String directoryPath = @"D:\Northwoods\Bin\DG99 Reports\";
        static readonly String textFile = @" Report.txt";


        static void Main(string[] args)
        {
            using (sqlConnection)
            {
                StringBuilder errorMessages = new StringBuilder();

                try
                {
                    try
                    {
                        sqlConnection.Open();

                        SqlCommand getNumberOfBadDocuments = new SqlCommand(
                            "Use OnBase " +
                            "select distinct count(id.itemnum) " +
                            "from hsi.itemdata id inner " +
                            "join " +
                            "hsi.itemdatapage idp on id.itemnum = idp.itemnum " +
                            "where idp.diskgroupnum = 99 "
                            , sqlConnection);

                        int totalBadDocuments = (int)getNumberOfBadDocuments.ExecuteScalar();

                        String[] badDocumentHandles = new String[totalBadDocuments];
                                                
                        badDocumentHandles = GetListOfBadDGDocuments(totalBadDocuments);

                        int i = 0;
                        
                        if(Directory.Exists(directoryPath))
                        {
                            using StreamWriter file = new StreamWriter(directoryPath + hostName + textFile);

                            file.WriteLine("------------ OnBase Diskgroup Scanner " + today + " ------------");
                            file.WriteLine();
                            file.WriteLine();


                            sqlConnection.Close();

                            if (totalBadDocuments > 0)
                            {
                                while (i < totalBadDocuments)
                                {
                                    file.WriteLine(badDocumentHandles[i]);
                                    i++;
                                }

                                file.WriteLine();
                                file.WriteLine();
                                file.WriteLine("------------ OnBase Diskgroup Scanner " + today + " ------------");
                            }
                        }
                        else
                        {
                            Directory.CreateDirectory(directoryPath);
                            using StreamWriter file = new StreamWriter(directoryPath + hostName + textFile);

                            file.WriteLine("------------ OnBase Diskgroup Scanner " + today + " ------------");
                            file.WriteLine();
                            file.WriteLine();


                            sqlConnection.Close();

                            if (totalBadDocuments > 0)
                            {
                                while (i < totalBadDocuments + 1)
                                {
                                    file.WriteLine(badDocumentHandles[i]);
                                    i++;
                                }

                                file.WriteLine();
                                file.WriteLine();
                                file.WriteLine("------------ OnBase Diskgroup Scanner " + today + " ------------");
                            }

                        }

                        if (totalBadDocuments > 0)
                        {
                            SendEmail(email);
                        }
                    }


                    
                    catch (SqlException sqlEx)
                    {
                        for (int s = 0; s < sqlEx.Errors.Count; s++)
                        {
                            errorMessages.Append("Index #" + s + "\n" +
                                "Message: " + sqlEx.Errors[s].Message + "\n" +
                                "LineNumber: " + sqlEx.Errors[s].LineNumber + "\n" +
                                "Source: " + sqlEx.Errors[s].Source + "\n" +
                                "Procedure: " + sqlEx.Errors[s].Procedure + "\n");
                        }
                        Console.WriteLine(errorMessages.ToString());
                        SendEmail(email, errorMessages.ToString());
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    SendEmail(email, ex.Message);
                }
            }
        }

        static void SendEmail(string email)
        {

            string to = email;
            string from = "dhhrrapidshelpdesk@wv.gov";
            MailMessage message = new MailMessage(from, to);
            message.Subject = @"" + hostName + " OnBase Diskgroup Scanner";
            message.Body = @"" + hostName + "'s report is attached.";
            SmtpClient client = new SmtpClient("10.200.146.27", 25);
            Attachment report;
            report = new Attachment(directoryPath + hostName + textFile);
            message.Attachments.Add(report);

            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught: ", ex.ToString());
            }
        }

        static void SendEmail(string email, string errorMessage)
        {

            string to = email;
            string from = "dhhrrapidshelpdesk@wv.gov";
            MailMessage message = new MailMessage(from, to);
            message.Subject = "Exception When Running " + hostName + "'s DG99 Finder";
            message.Body = @"The DG99 Finder did not run successfully on " + hostName + ". The exception caught was: " + errorMessage;
            SmtpClient client = new SmtpClient("10.200.146.27", 25);

            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught: ", ex.ToString());
            }

        }

        static String[] GetListOfBadDGDocuments(int numberOfBadDocs)
        {
            SqlCommand getDocuments = new SqlCommand(
            "Use OnBase " +
            "select distinct id.itemnum, id.itemname, idp.diskgroupnum " +
            "from hsi.itemdata id inner join " +
            "hsi.itemdatapage idp on id.itemnum = idp.itemnum " +
            "where idp.diskgroupnum = 99 "
            , sqlConnection);

            String[] badDocumentArray = new String[numberOfBadDocs];
            Int64 i = 0;

            using (SqlDataReader reader = getDocuments.ExecuteReader())
            {

                while (reader.Read())
                {
                  badDocumentArray[i] = reader["itemnum"].ToString();
                  i++;
                }
            }

            return badDocumentArray;
        }
    }
}
