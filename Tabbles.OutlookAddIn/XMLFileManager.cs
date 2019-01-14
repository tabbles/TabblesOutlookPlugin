using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Xml;


namespace Tabbles.OutlookAddIn
{
    public class XMLFileManager
    {



        private readonly string fileName = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +  @"\Tabbles.xml";
        private readonly string syncVal = "SyncValue";

       
        private const string ValueSyncPerformed = "SyncPerformed";
        private const string ValueDontAskForSync = "DontAskForSync";


        public XMLFileManager()
        {
            
        }

    
        private bool syncValue;

        private bool SyncStatus
        {
            get { return syncValue; }

            set 
            {
                if (syncValue != value)
                {
                    XmlDocument _xmlDoc = new XmlDocument();
                    _xmlDoc.Load(fileName);
                    XmlElement _xmlElem = _xmlDoc.DocumentElement;
                    _xmlElem.InnerText = value.ToString();
                    syncValue = Convert.ToBoolean(_xmlElem.InnerText);
                    _xmlDoc.Save(fileName);
                }
            }
        }
        

        public void CreateSettingsFile()
        {

            if (System.IO.File.Exists(fileName) == false)
            {

                XmlDocument _xmlDoc = new XmlDocument();
                XmlElement _xmlElem = _xmlDoc.CreateElement(syncVal);
                _xmlDoc.AppendChild(_xmlElem);
                _xmlElem.InnerText = "false";

               
                _xmlDoc.Save(fileName);
                
            }

        }

      
        public void SetSyncPerformed(bool performed)
        {
            //return SyncStatus;

            SyncStatus = performed;
        }

        public bool IsSyncPerformed()
        {
            return SyncStatus;
        }

        public void SetDontAskForSync(bool bDontAsk)
        {
          //  WriteBooleanValue(ValueDontAskForSync, dontAsk);
            SyncStatus = bDontAsk;
        }

        public bool IsDontAskForSync()
        {
           // return ReadBooleanValue(ValueDontAskForSync);
            return SyncStatus;
        }

        //private bool ReadBooleanValue(string valueName)
        //{
        //    try
        //    {
        //        RegistryKey baseKey = Registry.CurrentUser.OpenSubKey(MainSubKey, false);
        //        if (baseKey != null && baseKey.GetValue(valueName) != null &&
        //            baseKey.GetValueKind(valueName) == RegistryValueKind.DWord)
        //        {
        //            int result = (int)baseKey.GetValue(valueName);
        //            return result != 0;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Logger.Log(ex.ToString());
        //        //will return false
        //    }

        //    return false;
        //}

        //private void WriteBooleanValue(string valueName, bool value)
        //{
        //    try
        //    {
        //        RegistryKey baseKey = Registry.CurrentUser.OpenSubKey(MainSubKey, true);
        //        if (baseKey != null)
        //        {
        //            baseKey.SetValue(valueName, value ? 1 : 0, RegistryValueKind.DWord);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Logger.Log(ex.ToString());
        //    }
        //}
    

    



    }

}
