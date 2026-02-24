using System;
using AudioSwitcher.AudioApi;

namespace FortyOne.AudioSwitcher.Helpers
{
    public static class VolumeHelper
    {
        public static double Get()
        {
            var device = AudioDeviceManager.Controller.DefaultPlaybackDevice;
            if (device != null)
            {
                return device.Volume;
            }
            return 0;
        }

        public static async System.Threading.Tasks.Task Set(double volume)
        {
            var device = AudioDeviceManager.Controller.DefaultPlaybackDevice;
            if (device != null)
            {
                await device.SetVolumeAsync(volume);
                if (volume < 1)
                {
                    await device.SetMuteAsync(true);
                }
                else if (device.IsMuted)
                {
                    await device.SetMuteAsync(false);
                }
            }
        }
    }
}
