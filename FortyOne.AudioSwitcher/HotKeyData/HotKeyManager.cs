using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using FortyOne.AudioSwitcher.Configuration;

namespace FortyOne.AudioSwitcher.HotKeyData
{
    public static class HotKeyManager
    {
        private static readonly List<HotKey> _hotkeys = new List<HotKey>();
        public static BindingList<HotKey> HotKeys = new BindingList<HotKey>();
        public static HotKey QuickSwitchHotKey { get; set; }

        static HotKeyManager()
        {
            LoadHotKeys();
            RefreshHotkeys();
        }

        public static event EventHandler HotKeyPressed;

        public static void ClearAll()
        {
            foreach (var hk in _hotkeys)
            {
                hk.UnregisterHotkey();
            }

            Program.Settings.HotKeys = "";

            if (QuickSwitchHotKey != null)
            {
                QuickSwitchHotKey.UnregisterHotkey();
                QuickSwitchHotKey = null;
                Program.Settings.QuickSwitchHotKey = "";
            }

            LoadHotKeys();
            RefreshHotkeys();
        }

        public static void LoadHotKeys()
        {
            try
            {
                foreach (var hk in _hotkeys)
                {
                    hk.UnregisterHotkey();
                }

                _hotkeys.Clear();

                var hotkeydata = Program.Settings.HotKeys;
                LoadQuickSwitchHotKey();

                if (string.IsNullOrEmpty(hotkeydata))
                {
                    System.Diagnostics.Debug.WriteLine("No hotkey data found");
                    RefreshHotkeys();
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"Loading hotkeys from: {hotkeydata}");

                var entries = hotkeydata.Split(new[] { ",", "[", "]" }, StringSplitOptions.RemoveEmptyEntries);

                for (var i = 0; i < entries.Length; i++)
                {
                    try
                    {
                        var key = int.Parse(entries[i++]);
                        var modifiers = int.Parse(entries[i++]);
                        var hk = new HotKey();

                        var r = new Regex(ConfigurationSettings.GUID_REGEX);
                        var matches = r.Matches(entries[i]);
                        if (matches.Count == 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"No GUID match found in: {entries[i]}");
                            continue;
                        }
                        hk.DeviceId = new Guid(matches[0].ToString());

                        hk.Modifiers = (Modifiers)modifiers;
                        hk.Key = (Keys)key;
                        _hotkeys.Add(hk);
                        hk.HotKeyPressed += hk_HotKeyPressed;
                        hk.RegisterHotkey();
                        
                        System.Diagnostics.Debug.WriteLine($"Loaded hotkey: Key={hk.Key}, Modifiers={hk.Modifiers}, DeviceId={hk.DeviceId}");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error loading hotkey entry: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Critical error loading hotkeys: {ex.Message}");
                // Don't clear settings on error
                //Program.Settings.HotKeys = "";
            }
        }

        private static void hk_HotKeyPressed(object sender, EventArgs e)
        {
            if (HotKeyPressed != null)
                HotKeyPressed(sender, e);
        }

        public static void LoadQuickSwitchHotKey()
        {
            try
            {
                if (QuickSwitchHotKey != null)
                {
                    QuickSwitchHotKey.UnregisterHotkey();
                    QuickSwitchHotKey = null;
                }

                var qsData = Program.Settings.QuickSwitchHotKey;
                if (string.IsNullOrEmpty(qsData))
                    return;

                var entries = qsData.Split(new[] { ",", "[", "]" }, StringSplitOptions.RemoveEmptyEntries);

                if (entries.Length >= 2)
                {
                    var key = int.Parse(entries[0]);
                    var modifiers = int.Parse(entries[1]);

                    var hk = new HotKey();
                    hk.Modifiers = (Modifiers)modifiers;
                    hk.Key = (Keys)key;
                    hk.DeviceId = Guid.Empty; // Not tied to a specific device

                    QuickSwitchHotKey = hk;
                    hk.HotKeyPressed += hk_HotKeyPressed;
                    hk.RegisterHotkey();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading Quick Switch hotkey: {ex.Message}");
            }
        }

        public static void SaveQuickSwitchHotKey()
        {
            try
            {
                if (QuickSwitchHotKey == null || QuickSwitchHotKey.Key == Keys.None)
                {
                    Program.Settings.QuickSwitchHotKey = "";
                    return;
                }

                var hotkeydata = "[" + (int)QuickSwitchHotKey.Key + "," + (int)QuickSwitchHotKey.Modifiers + "]";
                Program.Settings.QuickSwitchHotKey = hotkeydata;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving Quick Switch hotkey: {ex.Message}");
            }
        }

        public static bool SetQuickSwitchHotKey(HotKey hk)
        {
            if (DuplicateHotKey(hk))
                return false;

            if (QuickSwitchHotKey != null)
            {
                QuickSwitchHotKey.UnregisterHotkey();
            }

            QuickSwitchHotKey = hk;
            QuickSwitchHotKey.HotKeyPressed += hk_HotKeyPressed;
            QuickSwitchHotKey.RegisterHotkey();
            
            SaveQuickSwitchHotKey();
            return true;
        }
        public static void SaveHotKeys()
        {
            try 
            {
                var hotkeydata = "";
                foreach (var hk in _hotkeys)
                {
                    // Validate the hotkey before saving
                    if (hk.Key == Keys.None || hk.DeviceId == Guid.Empty)
                    {
                        System.Diagnostics.Debug.WriteLine($"Skipping invalid hotkey: Key={hk.Key}, DeviceId={hk.DeviceId}");
                        continue;
                    }
                    
                    hotkeydata += "[" + (int)hk.Key + "," + (int)hk.Modifiers + "," + hk.DeviceId + "]";
                }
                
                System.Diagnostics.Debug.WriteLine($"Saving hotkeys: {hotkeydata}");
                
                // Only save if we have data to save
                if (string.IsNullOrEmpty(hotkeydata))
                    hotkeydata = "[]"; // Save empty array instead of empty string
                    
                Program.Settings.HotKeys = hotkeydata;
                
                // Verify the save worked by reading it back
                var savedData = Program.Settings.HotKeys;
                if (savedData != hotkeydata)
                {
                    System.Diagnostics.Debug.WriteLine($"Warning: Saved hotkey data doesn't match. Expected: {hotkeydata}, Got: {savedData}");
                }
                
                RefreshHotkeys();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving hotkeys: {ex.Message}");
                throw; // Re-throw to notify the UI layer
            }
        }

        public static bool AddHotKey(HotKey hk)
        {
            //Check that there is no duplicate
            if (DuplicateHotKey(hk))
                return false;

            hk.HotKeyPressed += hk_HotKeyPressed;
            hk.RegisterHotkey();

            if (!hk.IsRegistered)
                return false;

            _hotkeys.Add(hk);

            SaveHotKeys();
            return true;
        }

        public static void RefreshHotkeys()
        {
            HotKeys.Clear();
            var filterInvalid = !Program.Settings.ShowUnknownDevicesInHotkeyList;
            IEnumerable<HotKey> hotkeyList = _hotkeys;
            if (filterInvalid)
                hotkeyList = hotkeyList.Where(x => x.Device != null);
            
            foreach (var k in hotkeyList)
            {
                HotKeys.Add(k);
            }
        }

        public static bool DuplicateHotKey(HotKey hk)
        {
            if (QuickSwitchHotKey != null && hk.Key == QuickSwitchHotKey.Key && hk.Modifiers == QuickSwitchHotKey.Modifiers)
                return true;

            return _hotkeys.Any(k => hk.Key == k.Key && hk.Modifiers == k.Modifiers);
        }

        public static void DeleteHotKey(HotKey hk)
        {
            //Ensure its unregistered
            hk.UnregisterHotkey();
            _hotkeys.Remove(hk);
            SaveHotKeys();
        }
    }
}