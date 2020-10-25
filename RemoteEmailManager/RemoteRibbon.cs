using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.IO;

namespace RemoteEmailManager
{
    public partial class RemoteRibbon
    {
        
        private void RemoteRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnCopy_Click(object sender, RibbonControlEventArgs e)
        {
            SetClipboard();
        }


        private void SetClipboard()
        {
            string expMessage = "Your current selected email ";
            var folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RemoteCopy");

            //clear folder.
            System.IO.DirectoryInfo di = new DirectoryInfo(folderPath);
            foreach (FileInfo file in di.EnumerateFiles())
            {
                file.Delete();
            }

            try
            {
                if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
                {
                    foreach (var selObject in Globals.ThisAddIn.Application.ActiveExplorer().Selection)
                    {
                        if (selObject is Outlook.MailItem)
                        {   
                            Outlook.MailItem mailItem =
                                (selObject as Outlook.MailItem);
                            string emailPath = getEmailPath(mailItem);
                            mailItem.SaveAs(emailPath);                                                  
                        }

                    }
                }
                //set all the items to clipboard
                
                CopyFile(folderPath);
            }
            catch (Exception ex)
            {
                expMessage = ex.Message;
                MessageBox.Show(expMessage);
            }
            
        }

        private string getEmailPath(Outlook.MailItem emailItem)
        {
            var folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),"RemoteCopy");

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fName = Path.Combine(folderPath, emailItem.Subject + ".msg");
            if (File.Exists(fName))
            {
                File.Delete(fName);
            }
            return fName ;
        }

        private void CopyFile(string folderPath)
        {
            System.Collections.Specialized.StringCollection FileCollection = new System.Collections.Specialized.StringCollection();
            string[] fileEntries = Directory.GetFiles(folderPath);
            foreach (string FileToCopy in fileEntries)
            {
                FileCollection.Add(FileToCopy);
            }

            Clipboard.SetFileDropList(FileCollection);
        }
    }
}
