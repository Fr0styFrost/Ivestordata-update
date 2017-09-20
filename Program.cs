using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using EAGetMail; //add EAGetMail namespace
using System.Threading;
using System.Runtime.InteropServices;
using System.Data;


namespace DailyData
{
    class Program
    {      
       static void getemailattachments()
       {
           // Create a folder named inbox under current directory
           // to save the email retrieved.
           string curpath = Directory.GetCurrentDirectory();
           string mailbox = String.Format("{0}\\inbox", curpath);

           // If the folder is not existed, create it.
           if (!Directory.Exists(mailbox))
           {
               Directory.CreateDirectory(mailbox);
           }

           // Gmail IMAP4 server is "imap.gmail.com"
           MailServer oServer = new MailServer("imap.gmail.com", "rcisinvestordata@gmail.com", "RCISkey1", ServerProtocol.Imap4);
           MailClient oClient = new MailClient("TryIt");         
           oServer.SSLConnection = true;   // Set SSL connection,        
           oServer.Port = 993;             // Set 993 IMAP4 port

           try
           {
               oClient.Connect(oServer);
               MailInfo[] infos = oClient.GetMailInfos();


               for (int i = 0; i < infos.Length; i++)
               {
                   MailInfo info = infos[i];
                   Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}", info.Index, info.Size, info.UIDL);

                   // Download email from GMail IMAP4 server
                   Mail oMail = oClient.GetMail(info);
                   oMail.DecodeTNEF();
                   Attachment[] atts = oMail.Attachments;
                   int count = atts.Length;
                   string tempFolder = "c:\\temp";
                   if (!Directory.Exists(tempFolder))
                   {
                       Directory.CreateDirectory(tempFolder);
                   }
                   //Saves attachments.
                   for (int c = 0; c < count; c++)
                   {
                       Attachment att = atts[c];
                       string attname = String.Format("{0}\\{1}", tempFolder, att.Name); att.SaveAs(attname, true);
                   }
                   // Mark email as deleted in GMail account to prevent duplicats.
                   oClient.Delete(info);
               }

               // Quit and purge emails marked as deleted from Gmail IMAP4 server.
               oClient.Quit();

           }

           catch (Exception ep)
           {
               Console.WriteLine(ep.Message);
               Console.ReadKey();
           }
       }    //Method that gets emails and saves attachments to temp folder
       static void Sortattachments() 
       {
           DirectoryInfo di = new DirectoryInfo(@"c:\temp\");
           foreach (FileInfo file in di.GetFiles("*.zip"))
           {
               int delete = 1; // variable to indicate if file is safe to be deleted.
               try   //try & catch to attempt to copy files to a network drive location and indicate where the error may be.
               {                                     
                   File.Copy(@"c:\temp\" + file.Name, @"Z:\" + file.Name,true); 
                   Console.WriteLine(file.Name + " has been coppied to Z:\\");
               }
               catch
               {
                   Console.WriteLine("Cannot Access Z: drive Main work station my be turned off.");
                   delete = 0;
               }
               try   //try & catch to attempt to copy files to a network drive location and indicate where the error may be.
               {
                   File.Copy(@"c:\temp\" + file.Name, @"X:\" + file.Name, true);
                   Console.WriteLine(file.Name + " has been coppied to X:\\");
               }
               catch
               {
                   Console.WriteLine("Cannot Access X: drive Dale my be turned off.");
                   delete = 0;
               }
               try    //try & catch to attempt to copy files to a network drive location and indicate where the error may be.
               {
                   File.Copy(@"c:\temp\" + file.Name, @"Y:\" + file.Name, true);
                   Console.WriteLine(file.Name + " has been coppied to y:\\");
               }
               catch
               {
                   Console.WriteLine("Cannot Access Y: drive Gizmo my be turned off.");
                   delete = 0;

               }
               try    //try & catch to attempt to copy files to a network drive location and indicate where the error may be.
               {
                   File.Copy(@"c:\temp\" + file.Name, @"W:\" + file.Name, true);
                   Console.WriteLine(file.Name + " has been coppied to W:\\");
               }
               catch
               {
                   Console.WriteLine("Cannot Access W: drive Richard XP more my be turned off.");
                   delete = 0;
               }

               if (delete == 1)  //if no errors where found file is safe to be deleted.                                                                 //
               {
                   Console.WriteLine(file.Name + " has been coppied successfully");
                   file.Delete();
               }
               else //if errors are found give a warning and wait 5 min before trying again.
               {                   
                   Console.WriteLine("There was an error copping the files waiting 5 min");
                   Thread.Sleep(1000 * 60 * 5);  // timer in ms to wait 5 min before restarting the whole process.                  

               }               
           }
       }        //Method moves all files to correct dirrectory or moves it to storage.            
       static void repeat()
       {
           DateTime startdate = DateTime.Today.AddHours(08.5);
           DateTime enddate = DateTime.Today.AddHours(13.5);       
           if (DateTime.Now > startdate && DateTime.Now < enddate)
           {
               Sortattachments(); //calls the sort method               
           }
           else
           {
               getemailattachments(); //calls the get mail method             
           }
       }                //Method that allows to cycle through methods during the day.                               

       static void Main(string[] args)
        {
            getemailattachments(); //calls the get mail method on startup to get mails on pc next startup
            while (true)
            {
                repeat();
            }
         }                
    }
}



            




        