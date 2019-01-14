using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace Tabbles.OutlookAddIn
{
    public static class RegistryManager
    {
        private const string MainSubKey = @"Software\Yellow Blue Soft\Tabbles for Outlook";
        private const string ValueSyncPerformed = "SyncPerformed";
        private const string ValueDontAskForSync = "DontAskForSync";

        public static void SetSyncPerformed(bool performed)
        {
            WriteBooleanValue(ValueSyncPerformed, performed);
        }

        public static bool IsSyncPerformed()
        {
            return ReadBooleanValue(ValueSyncPerformed);
        }

        public static void SetDontAskForSync(bool dontAsk)
        {
            WriteBooleanValue(ValueDontAskForSync, dontAsk);
        }

        public static bool IsDontAskForSync()
        {
            return ReadBooleanValue(ValueDontAskForSync);
        }

        private static bool ReadBooleanValue(string valueName)
        {
            try
            {
                var baseKey = Registry.CurrentUser.OpenSubKey(MainSubKey, false);
                if (baseKey != null && baseKey.GetValue(valueName) != null &&
                    baseKey.GetValueKind(valueName) == RegistryValueKind.DWord)
                {
                    var result = (int)baseKey.GetValue(valueName);
                    return result != 0;
                }
            }
            catch (Exception ex)
            {
                Log.log(ex.ToString());
                //will return false
            }

            return false;
        }

        private static void WriteBooleanValue(string valueName, bool value)
        {
            try
            {
                RegistryKey baseKey = Registry.CurrentUser.OpenSubKey(MainSubKey, true);
                if (baseKey != null)
                {
                    baseKey.SetValue(valueName, value ? 1 : 0, RegistryValueKind.DWord);
                }
            }
            catch (Exception ex)
            {
                Log.log(ex.ToString());
            }
        }
    }
}
