using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using FortyOne.AudioSwitcher.Configuration;

namespace FortyOne.AudioSwitcher.Helpers
{
    public class VolumeHook : IDisposable
    {
        private const int STEP = 1;
        private IntPtr _hookId = IntPtr.Zero;
        private readonly LowLevelKeyboardProc _proc;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        public VolumeHook()
        {
            _proc = HookCallback;
            Hook();
        }

        public void Hook()
        {
            if (_hookId != IntPtr.Zero) return;

            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                _hookId = SetWindowsHookEx(13, _proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        public void Unhook()
        {
            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)0x0100) // WM_KEYDOWN
            {
                int vk = Marshal.ReadInt32(lParam);

                if (vk == 0xAF) // Volume Up
                {
                    System.Threading.Tasks.Task.Run(async () =>
                    {
                        var newVol = Math.Min(VolumeHelper.Get() + STEP, 100);
                        await VolumeHelper.Set(newVol);
                        if (Program.Settings.ShowVolumeOSD)
                            AudioSwitcher.Instance.BeginInvoke((Action)(() => VolumeOSD.ShowVolume((int)newVol, VolumeHelper.IsMuted())));
                    });
                    return (IntPtr)1;
                }
                if (vk == 0xAE) // Volume Down
                {
                    System.Threading.Tasks.Task.Run(async () =>
                    {
                        var newVol = Math.Max(VolumeHelper.Get() - STEP, 0);
                        await VolumeHelper.Set(newVol);
                        if (Program.Settings.ShowVolumeOSD)
                            AudioSwitcher.Instance.BeginInvoke((Action)(() => VolumeOSD.ShowVolume((int)newVol, VolumeHelper.IsMuted())));
                    });
                    return (IntPtr)1;
                }
                if (vk == 0xAD) // Volume Mute
                {
                    System.Threading.Tasks.Task.Run(async () =>
                    {
                        await VolumeHelper.ToggleMute();
                        var currentVol = VolumeHelper.Get();
                        if (Program.Settings.ShowVolumeOSD)
                            AudioSwitcher.Instance.BeginInvoke((Action)(() => VolumeOSD.ShowVolume((int)currentVol, VolumeHelper.IsMuted())));
                    });
                    return (IntPtr)1;
                }
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        public void Dispose()
        {
            Unhook();
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
    }
}
