using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Controls;

namespace MailChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            string inDir = ConfigurationManager.AppSettings.Get("INdir");
            string outDir = ConfigurationManager.AppSettings.Get("OUTdir");         

            DataTable reqRes;
            string f;
            List<string> folderList = new List<string>();
            folderList = ListFolders();
            SearchQuery query;
            query = SearchQuery.NotSeen.And(SearchQuery.SubjectContains("Портал Росреестра: заявление выполн"));
            SearchQuery notSeen = SearchQuery.NotSeen;
            //query = SearchQuery.FromContains("portal@rosreestr.ru");            

            #region Для отлалдки
            //Console.WriteLine("\n Выберите нужную папку:\n");
            //string[] folders = new string[folderList.Count()];
            //int i = 0;
            //foreach (var folder in folderList)
            //{
            //    Console.WriteLine("[{0}] {1}", i, folder);
            //    folders[i] = folder;
            //    i++;
            //}
            //f = Console.ReadLine();


            ////ListMessages(folders[System.Convert.ToInt32(f)], query);
            //reqRes = RequestParser(folders[System.Convert.ToInt32(f)], query);
            #endregion

            reqRes = RequestParser(ConfigurationManager.AppSettings.Get("FOLDER"), query);
            //CSVUtility.CSVUtility.ToCSV(reqRes, @ConfigurationManager.AppSettings.Get("OUTdir") + "out.csv");
            
            CheckRequestFromDir(inDir, outDir, reqRes);
            CSVUtility.CSVUtility.ToXLSX(reqRes, @ConfigurationManager.AppSettings.Get("OUTdir") + "out.xlsx", @ConfigurationManager.AppSettings.Get("TEMPL_PATH"));
            Console.WriteLine("Готово!");
            Console.ReadKey();
        }


        public static List<string> ListFolders()
        {
            List<string> folderList = new List<string>();

            using (var client = new ImapClient())
            {
                client.Connect(ConfigurationManager.AppSettings.Get("IMAPserver"), 993, SecureSocketOptions.SslOnConnect);

                client.Authenticate(ConfigurationManager.AppSettings.Get("LOGIN"), ConfigurationManager.AppSettings.Get("PASS"));

                var personal = client.GetFolders(client.PersonalNamespaces[0]);
                foreach (var folder in personal)
                {
                    folderList.Add(folder.FullName);
                    folder.Status(StatusItems.Unread);
                    Console.WriteLine("[folder] {0} Сообщений: {1}", folder.Name, folder.Unread);
                }

                client.Disconnect(true);
                return folderList;
            }
        }