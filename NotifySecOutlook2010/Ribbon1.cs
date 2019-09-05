
using NotifySecOutlook2010.Properties;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;


namespace NotifySecOutlook2010
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("NotifySecOutlook2010.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public string GetButtonLabel(Office.IRibbonControl control)
        {
            return "דווח על מייל חשוד";
        }

        public Bitmap GetImage(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.shield);
        }

        public void ShowMessageClick(Office.IRibbonControl control)
        {

            if (control.Id == "button1D" || control.Id == "button1D2")
            {
                CreateNewMailToSecurityTeam(control);
                //System.Windows.Forms.MessageBox.Show("Button clicked!");

            }
        }

        private void CreateNewMailToSecurityTeam(IRibbonControl control)
        {

            Selection selection =
                Globals.ThisAddIn.Application.ActiveExplorer().Selection;

            if (selection.Count == 1)   // Check that selection is not empty.
            {
                object selectedItem = selection[1];   // Index is one-based.
                Object mailItemObj = selectedItem as Object;
                MailItem mailItem = null;// selectedItem as MailItem;
                if (selection[1] is Outlook.MailItem)
                {
                    mailItem = selectedItem as MailItem;
                }

                MailItem tosend = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                tosend.Attachments.Add(mailItemObj);
                tosend.BodyFormat = OlBodyFormat.olFormatHTML; //html for rtl

                #region create mail from default
                try
                {

                    tosend.To = "YOUR_SECURITY_TEAM@MAIL.CO.IL"; // >>>>>>>> enter security team mail here.
                    tosend.CC = "Help_DESK@MAIL.CO.IL"; // >>>>>>>> enter help desk team mail here (e.g SYSAID).
                    tosend.Subject = "דיווח על מייל חשוד";

                    // fix ltr
                    tosend.Body = "<b>שלום";
                    tosend.Body += "<br/>";
                    tosend.Body += "מצורף בזה מייל שנראה לי חשוד";
                    tosend.Body += "\n";
                    tosend.Body += "אודה לבדיקתכם.";
                    tosend.Body += "</b><br/>";
                    tosend.Body += "<br/>";

                    #endregion
                    tosend.Body += GetCurrentUserInfos();
                    //add message headers
                    Outlook.PropertyAccessor olPA = mailItem.PropertyAccessor;
                    String Header = olPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E");
                    tosend.Body += "<br/><br/> <b> Message headers: </b><br/>" + Header + "<br/>";
                    //rtl
                    tosend.HTMLBody = "<HTML><BODY><div style='direction:rtl;text-align:right'>" + tosend.Body + "</div></BODY></HTML>";
                    tosend.Save();
                    tosend.Display();
                    
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Using default template" + ex.Message);

                    MailItem mi = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                    mi.To = "help_desk@mail.co.il"; //cant send mail object (e.g user try to send calender item).
                    mi.Subject = "Security addin error";
                    String txt = ("פרטים נוספים לגבי תקלה בהפעלת התוסף: " + ex);
                    mi.Body = txt;
                    mi.Save();
                    mi.Display();
                }
            }
            else if (selection.Count < 1)   // Check that selection is not empty.
            {
                MessageBox.Show("אנא בחר את ההודעה החשודה ");
            }
            else if (selection.Count > 1)
            {
                MessageBox.Show("אנא בחר את ההודעה החשודה -  יש לבחור הודעה אחת בלבד");
            }
            else
            {
                MessageBox.Show("נראה שמשהו השתבש...");
            }


        }
        #endregion


        //add reporting user information
        public String GetCurrentUserInfos()
        {

            String wComputername = System.Environment.MachineName + " (" + System.Environment.OSVersion.ToString() + ")";
            String wUsername = System.Environment.UserDomainName + "\\" + System.Environment.UserName;

            string str = "<b>מידע נוסף:</b>";

            
            Outlook.AddressEntry addrEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
            if (addrEntry.Type == "EX")
            {
                Outlook.ExchangeUser currentUser =
                    Globals.ThisAddIn.Application.Session.CurrentUser.
                    AddressEntry.GetExchangeUser();
                if (currentUser != null)
                {
                    str += " שם עובד מדווח: " + currentUser.Name;
                    str += " STMP address: " + currentUser.PrimarySmtpAddress;
                    str += " תפקיד: " + currentUser.JobTitle;
                    str += " מחלקה: " + currentUser.Department;
                    str += " מיקום: " + currentUser.OfficeLocation;
                    str += " טלפון: " + currentUser.BusinessTelephoneNumber;
                    str += " נייד: " + currentUser.MobileTelephoneNumber;
                    str += "<br/>"; 


                }
            }
            str += "\n - Windows username:" + wUsername;
            str += "\n - Computername:" + wComputername;
            str += "<br/>";
            return str;
        }
        




        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
       

    }
}