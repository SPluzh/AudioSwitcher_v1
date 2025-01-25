using System;
using System.Collections.Generic;
using System.IO;
using fastJSON;

namespace FortyOne.AudioSwitcher.Configuration
{
    public class JsonSettings : ISettingsSource
    {
        private readonly object _mutex = new object();
        private string _path;
        private IDictionary<string, string> _settingsObject;

        public JsonSettings()
        {
            _settingsObject = new Dictionary<string, string>();
        }

        public void SetFilePath(string path)
        {
            _path = path;
        }

        public void Load()
        {
            lock (_mutex)
            {
                try
                {
                    if (File.Exists(_path))
                        _settingsObject = JSON.ToObject<Dictionary<string, string>>(File.ReadAllText(_path));
                }
                catch
                {
                    _settingsObject = new Dictionary<string, string>();
                }
            }
        }

        public void Save()
        {
            lock (_mutex)
            {
                try
                {
                    var json = JSON.ToJSON(_settingsObject);
                    var beautified = JSON.Beautify(json);
                    
                    // Write to a temporary file first
                    var tempPath = _path + ".tmp";
                    File.WriteAllText(tempPath, beautified);
                    
                    // If the write succeeded, move the temp file to the real location
                    if (File.Exists(_path))
                        File.Delete(_path);
                        
                    File.Move(tempPath, _path);
                    
                    System.Diagnostics.Debug.WriteLine($"Settings saved successfully to {_path}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error saving settings: {ex.Message}");
                    throw; // Re-throw so the application knows about the failure
                }
            }
        }

        public string Get(string key)
        {
            lock (_mutex)
            {
                return _settingsObject[key];
            }
        }

        public void Set(string key, string value)
        {
            lock (_mutex)
            {
                try
                {
                    _settingsObject[key] = value;
                    Save();
                    System.Diagnostics.Debug.WriteLine($"Successfully set and saved {key}={value}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error setting {key}={value}: {ex.Message}");
                    throw;
                }
            }
        }
    }
}