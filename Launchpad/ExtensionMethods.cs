using System;
using System.Reflection;
using Midi;

namespace ExtensionMethods
{
    public static class ExtensionMethods
    {
        public static Types.LAUNCHPAD_TYPE type(this Midi.InputDevice device)
        {
            if (device.Name.StartsWith("Launchpad S", StringComparison.InvariantCultureIgnoreCase))
            {
                return Types.LAUNCHPAD_TYPE.S;
            } else if (device.Name.StartsWith("Launchpad Mk2", StringComparison.InvariantCultureIgnoreCase)) {
                return Types.LAUNCHPAD_TYPE.MK2;
            }
            return Types.LAUNCHPAD_TYPE.UNKNOWN_DEVICE_TYPE;
        }
        public static Types.LAUNCHPAD_TYPE type(this Midi.OutputDevice device)
        {
            if (device.Name.StartsWith("Launchpad S", StringComparison.InvariantCultureIgnoreCase))
            {
                return Types.LAUNCHPAD_TYPE.S;
            }
            else if (device.Name.StartsWith("Launchpad Mk2", StringComparison.InvariantCultureIgnoreCase))
            {
                return Types.LAUNCHPAD_TYPE.MK2;
            }
            return Types.LAUNCHPAD_TYPE.UNKNOWN_DEVICE_TYPE;
        }
        public static bool equals(this Midi.InputDevice device, Midi.InputDevice foreignDevice)
        {
            return device.Name == foreignDevice.Name;
        }
        public static bool equals(this Midi.OutputDevice device, Midi.OutputDevice foreignDevice)
        {
            return device.Name == foreignDevice.Name;
        }
    }
}