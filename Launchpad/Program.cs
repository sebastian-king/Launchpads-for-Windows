using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.Threading;
using Midi;
using WindowsInput; // InputSimulator
using NAudio;
using ExtensionMethods; // Adds type to identify Launchpads
using System.Runtime.InteropServices;
using System.Management;
using System.Diagnostics;

namespace Launchpad {
    class Program
    {
        public enum CtrlTypes
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT,
            CTRL_CLOSE_EVENT,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT
        }

        [DllImport("Kernel32")]
        public static extern bool SetConsoleCtrlHandler(HandlerRoutine Handler, bool Add);
        public delegate bool HandlerRoutine(CtrlTypes CtrlType);
        private enum StdHandle { Stdin = -10, Stdout = -11, Stderr = -12 };
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetStdHandle(StdHandle std);
        [DllImport("kernel32.dll")]
        private static extern bool CloseHandle(IntPtr hdl);

        public static Thread volume_visualisation;
        
        public static List<InputDevice> inputDevice = new List<InputDevice>();
        public static List<OutputDevice> outputDevice = new List<OutputDevice>();
        public static bool[] on = new bool[121];

        public static bool muted;
        public static double[] volumesLeft  = new double[4];
        public static double[] volumesRight = new double[4];

        public static int[] audio_colours = new int[8]{48,49,50,51,51,35,19,3};

        static void Main(string[] args)
        {
            Console.WriteLine("Launching Launchpad program ...");

            initialiseDevices();

            WqlEventQuery insertQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBHub'");
            ManagementEventWatcher insertWatcher = new ManagementEventWatcher(insertQuery);
            insertWatcher.EventArrived += new EventArrivedEventHandler(DeviceInsertedEvent);
            insertWatcher.Start();
            
            WqlEventQuery removeQuery = new WqlEventQuery("SELECT * FROM __InstanceDeletionEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBHub'");
            ManagementEventWatcher removeWatcher = new ManagementEventWatcher(removeQuery);
            removeWatcher.EventArrived += new EventArrivedEventHandler(DeviceRemovedEvent);
            removeWatcher.Start();
            
            NAudio.CoreAudioApi.MMDeviceEnumerator MMDE = new NAudio.CoreAudioApi.MMDeviceEnumerator();
            NAudio.CoreAudioApi.MMDevice device = MMDE.GetDefaultAudioEndpoint(NAudio.CoreAudioApi.DataFlow.Render, NAudio.CoreAudioApi.Role.Console);
            muted = device.AudioEndpointVolume.Mute;
            AudioEndpointVolume_OnVolumeNotification(new NAudio.CoreAudioApi.AudioVolumeNotificationData(new Guid(), device.AudioEndpointVolume.Mute, device.AudioEndpointVolume.MasterVolumeLevel, new float[device.AudioMeterInformation.PeakValues.Count])); // set VolLevels to be null first time
            device.AudioEndpointVolume.OnVolumeNotification += AudioEndpointVolume_OnVolumeNotification;
            
            SetConsoleCtrlHandler(new HandlerRoutine(ConsoleCtrlCheck), true);
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(OnProcessExit);

            /*
            while (!exit) // be a program
            {
                // uh
            }
            */

            volume_visualisation = new Thread(visualise_volume);
            volume_visualisation.IsBackground = true;
            volume_visualisation.Start();

            Console.WriteLine("Press 'l' for a list of currently attached devices");
            Console.WriteLine("Press 'esc' to exit.");

            ConsoleKeyInfo key = Console.ReadKey();

            while (key.Key != ConsoleKey.Escape)
            {
                Console.WriteLine("IDs: ");
                foreach (InputDevice inputdevice in inputDevice)
                {
                    Console.Write(inputdevice.Name + ", ");
                }
                Console.WriteLine();
                Console.WriteLine("ODs: ");
                foreach (InputDevice outputdevice in inputDevice)
                {
                    Console.Write(outputdevice.Name + ", ");
                }
                Console.WriteLine();
                key = Console.ReadKey();
            }

            _exit(Int32.MaxValue);

            return;
        }
        private static bool ConsoleCtrlCheck(CtrlTypes ctrlType)
        {
            switch (ctrlType)
            {
                case CtrlTypes.CTRL_CLOSE_EVENT:
                case CtrlTypes.CTRL_LOGOFF_EVENT:
                case CtrlTypes.CTRL_SHUTDOWN_EVENT:
                    Console.WriteLine("Program being closed!");
                    _exit();
                    break;
                default:
                    _exit();
                    break;
            }
            return true;
        }
        static void OnProcessExit(object sender, EventArgs e) // 2 second timeout
        {
            // i feel like this doesn't serve any purpose anymore.
            Console.WriteLine("I'm out of here");
        }

        public static void NoteOn(NoteOnMessage msg)
        {
            Console.Write("[" + msg.Device.Name + "]: ON -> " + (int)msg.Pitch + " -> " + msg.Velocity);
            Console.WriteLine(" == SET TO ON <" + on[(int)msg.Pitch] + ">");

            if (((int)msg.Pitch == 88 || (int)msg.Pitch == 39) && msg.Velocity == 127)
            {
                //InputSimulator.SimulateKeyPress(VirtualKeyCode.VOLUME_MUTE);
                NAudio.CoreAudioApi.MMDeviceEnumerator MMDE = new NAudio.CoreAudioApi.MMDeviceEnumerator();
                NAudio.CoreAudioApi.MMDevice device = MMDE.GetDefaultAudioEndpoint(NAudio.CoreAudioApi.DataFlow.Render, NAudio.CoreAudioApi.Role.Console);
                 if (device.AudioEndpointVolume.Mute == true)
                {
                    device.AudioEndpointVolume.Mute = false;
                }
                else
                {
                    device.AudioEndpointVolume.Mute = true;
                }
            }

        }
        public static void NoteOff(NoteOffMessage msg)
        {
            Console.Write("[" + msg.Device.Name + "]: OFF -> " + (int)msg.Pitch + " -> " + msg.Velocity);
            Console.WriteLine(" == SET TO OFF");
        }
        public static void SysEx(SysExMessage msg)
        {
            StringBuilder hex = new StringBuilder(msg.Data.Length * 2);
            foreach (byte b in msg.Data)
            {
                hex.AppendFormat("{0:x2}", b);
            }
            if ("f0002029021815f7" == hex.ToString())
            {
                Console.WriteLine("[" + msg.Device.Name + "]: device text has finished/completed loop.");
            } else {
                Console.WriteLine("[" + msg.Device.Name + "]: SysEx Message: \n" + hex.ToString());
            }
        }
        public static void ControlChange(ControlChangeMessage msg)
        {
            switch ((int)msg.Control)
            {
                case 0:
                    if (msg.Value == 3)
                    {
                        Console.WriteLine("[" + msg.Device.Name + "]: device text has finished/completed loop.");
                    }
                    break;
                case 104:
                    if (msg.Value == 127)
                    {
                        if (volume_visualisation.IsAlive)
                        {
                            Console.WriteLine("Disabling volume visualisation.");
                            volume_visualisation.Abort();
                            for (int i = 0; i < inputDevice.Count; i++)
                            {
                                if (outputDevice[i].type() == Types.LAUNCHPAD_TYPE.S)
                                {
                                    for (int z = 0; z < 8; z++)
                                    {
                                        for (int n = 0; n < 8; n++)
                                        {
                                            outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)((z + 112) - (n * 16)), 0);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("Enabling volume visualisation");
                            volume_visualisation = new Thread(visualise_volume);
                            volume_visualisation.IsBackground = true;
                            volume_visualisation.Start();
                        }
                    }
                    break;
                case 111:
                    if (msg.Value == 0)
                    {
                        for (int i = 0; i < outputDevice.Count; i++)
                        {
                            outputDevice[i].SendControlChange(Channel.Channel1, msg.Control, 15);
                        }
                        Console.WriteLine("PROGRAM EXITING.");
                        CloseHandle(GetStdHandle(StdHandle.Stdin));
                    }
                    break;
                default:
                    Console.WriteLine("[" + msg.Device.Name + "]: Control: " + msg.Control + ", Value: " + msg.Value);
                    if ((int)msg.Value == 127)
                    {
                        for (int i = 0; i < outputDevice.Count; i++)
                        {
                            outputDevice[i].SendControlChange(Channel.Channel1, msg.Control, 15);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < outputDevice.Count; i++)
                        {
                            outputDevice[i].SendControlChange(Channel.Channel1, msg.Control, 0);
                        }
                    }
                    break;
            }
        }
        public static Byte[] textMk2(string text, int colour, bool loop = true) {
            var b = new byte[text.Length + 10];
            b[0] = 0xF0;
            b[1] = 0x00;
            b[2] = 0x20;
            b[3] = 0x29;
            b[4] = 0x02;
            b[5] = 0x18; // SysEx header -- 6
            b[6] = 0x14; // text packet -- 1
            b[7] = (byte)colour; // colour -- 1
            b[8] = (byte)(loop == true ? 1 : 0); // loop, 1 = yes -- 1
            // 10 PACKETS OF OVERHEAD, including ending packet
            char[] a = text.ToCharArray();
            for (int i = 0; i < a.Length; i++) // handle timing changes, done by using \x000{1-7} inline
            {
                b[i + 9] = Convert.ToByte(a[i]);
            }
            b[b.Length - 1] = 0xF7;
            return b;
        }
        public static Byte[] textS(string text, int colour, bool loop = true)
        {
            var b = new byte[text.Length + 7];
            b[0] = 0xF0;
            b[1] = 0x00;
            b[2] = 0x20;
            b[3] = 0x29;
            b[4] = 0x09; // SysEx header -- 5
            b[5] = (byte)(colour + (loop == true ? 64 : 0)); // colour/loop -- 1
            // 7 PACKETS OF OVERHEAD, including ending packet
            char[] a = text.ToCharArray();
            for (int i = 0; i < a.Length; i++) // handle timing changes, done by using \x000{1-7} inline
            {
                b[i + 6] = Convert.ToByte(a[i]);
            }
            b[b.Length - 1] = 0xF7;
            return b;
        }
        public static void initialiseDevices()
        {
            initialiseDevices(true);
        }
        public static void initialiseDevices(bool init)
        {
            InputDevice.UpdateInstalledDevices();
            OutputDevice.UpdateInstalledDevices();
            if (init)
            {
                for (int i = 0; i < InputDevice.InstalledDevices.Count; i++)
                {
                    if (InputDevice.InstalledDevices[i].type() != Types.LAUNCHPAD_TYPE.UNKNOWN_DEVICE_TYPE) // Launchpad*
                    {
                        for (int n = 0; n < OutputDevice.InstalledDevices.Count; n++)
                        {
                            if (OutputDevice.InstalledDevices[n].Name == InputDevice.InstalledDevices[i].Name)
                            {
                                inputDevice.Add(InputDevice.InstalledDevices[i]);
                                outputDevice.Add(OutputDevice.InstalledDevices[n]);
                                Console.WriteLine("Using input/output device: \"" + inputDevice[inputDevice.Count - 1].Name + "\"");
                            }
                        }
                    }
                }
            }
            for (int i = 0; i < inputDevice.Count; i++)
            {
                if (!inputDevice[i].IsOpen && !inputDevice[i].IsReceiving)
                {
                    inputDevice[i].Open();
                    inputDevice[i].ControlChange += new InputDevice.ControlChangeHandler(ControlChange);
                    inputDevice[i].NoteOn += new InputDevice.NoteOnHandler(NoteOn);
                    inputDevice[i].NoteOff += new InputDevice.NoteOffHandler(NoteOff);
                    inputDevice[i].SysEx += new InputDevice.SysExHandler(SysEx);
                    inputDevice[i].StartReceiving(null, true); // Note events will be received in another thread, null means start clock at 0

                    Console.WriteLine("Initiaiised device #" + i + ", device is type: " + inputDevice[i].type());

                    outputDevice[i].Open();

                    if (inputDevice[i].type() == Types.LAUNCHPAD_TYPE.S)
                    {
                        outputDevice[i].SendControlChange(Channel.Channel1, 0, 0); // clear S
                    }

                    if (inputDevice[i].type() == Types.LAUNCHPAD_TYPE.MK2)
                    {
                        outputDevice[i].SendSysEx(new Byte[] { 0xF0, 0x00, 0x20, 0x29, 0x02, 0x18, 0x0E, 0x0, 0xF7 }); // clear Mk2
                    }

                    outputDevice[i].SendSysEx(new Byte[] { 0xF0, 0x7E, 0x7F, 0x06, 0x01, 0xF7 }); // device info for BOTH S & Mk2

                    if (inputDevice[i].type() == Types.LAUNCHPAD_TYPE.S)
                    {
                        outputDevice[i].SendSysEx(textS("\x06READY", 15, false));
                    }
                    else if (inputDevice[i].type() == Types.LAUNCHPAD_TYPE.MK2)
                    {
                        outputDevice[i].SendSysEx(textMk2("\x06READY", 72, false));
                    }
                }
            }
        }
        public static void uninitialiseDevices()
        {
            //Console.WriteLine("ID: " + String.Join(", ", inputDevice));
            //Console.WriteLine("OD: " + String.Join(", ", outputDevice));
            foreach (InputDevice inputdevice in inputDevice)
            {
                //try
                //{
                    inputdevice.StopReceiving();
                    inputdevice.RemoveAllEventHandlers();
                    inputdevice.Close();
                //} catch (DeviceException)
                //{

                //}
            }
            inputDevice.Clear();
            foreach (OutputDevice outputdevice in outputDevice)
            {
                //try
                //{
                    outputdevice.Close();
                //}
                //catch (DeviceException)
                //{

                //}
            }
            outputDevice.Clear();
        }
        static void DeviceInsertedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            //foreach (var property in instance.Properties)
            //{
                //Console.WriteLine(property.Name + " I= " + property.Value);
            //}
            Console.WriteLine("A USB device has been plugged in.");
            checkPluggedIn();
        }
        static void DeviceRemovedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            //foreach (var property in instance.Properties)
            //{
            //Console.WriteLine(property.Name + " I= " + property.Value);
            //}
            //Console.WriteLine("A USB device has been unplugged.");
            //checkPluggedIn();

            //ProcessStartInfo Info = new ProcessStartInfo();
            //Info.Arguments = "/C ping 127.0.0.1 -n 2 >nul 2>&1 && \"" + System.Environment.GetCommandLineArgs()[0] + "\"";
            //Info.FileName = "cmd.exe";
            //Process.Start(Info);

            _exit(true);
        }
        public static void checkPluggedIn() {
            try
            {
                InputDevice.UpdateInstalledDevices();
                OutputDevice.UpdateInstalledDevices();
                for (int i = 0; i < InputDevice.InstalledDevices.Count; i++)
                {
                    //Console.WriteLine("1");
                    if (InputDevice.InstalledDevices[i].type() != Types.LAUNCHPAD_TYPE.UNKNOWN_DEVICE_TYPE) // Launchpad*
                    {
                        //Console.WriteLine("2");
                        for (int n = 0; n < OutputDevice.InstalledDevices.Count; n++)
                        {
                            //Console.WriteLine("3");
                            if (OutputDevice.InstalledDevices[n].Name == InputDevice.InstalledDevices[i].Name)
                            {
                                //Console.WriteLine("4");
                                bool pluggedIn = false;
                                for (int q = 0; q < inputDevice.Count; q++)
                                {
                                    //
                                    if (inputDevice[q].equals(InputDevice.InstalledDevices[i]))
                                    {
                                        //Console.WriteLine(InputDevice.InstalledDevices[i].Name + " is already plugged in at exists at i:" + q + "(" + inputDevice[q].Name + ")");
                                        pluggedIn = true;
                                        break;
                                    }
                                }
                                for (int q = 0; q < outputDevice.Count; q++)
                                {
                                    //
                                    if (outputDevice[q].equals(OutputDevice.InstalledDevices[n]))
                                    {
                                        //Console.WriteLine(OutputDevice.InstalledDevices[n].Name + " is already plugged in at exists at o:" + q + "(" + inputDevice[q].Name + ")");
                                        pluggedIn = true;
                                        break;
                                    }
                                }
                                if (pluggedIn == false)
                                {
                                    inputDevice.Add(InputDevice.InstalledDevices[i]);
                                    outputDevice.Add(OutputDevice.InstalledDevices[n]);
                                    Console.WriteLine("Plugged in and attached input/output device: \"" + inputDevice[inputDevice.Count - 1].Name + "\"");
                                    initialiseDevices(false);
                                    return;
                                }
                            }
                        }
                    }
                }
            }
            catch (DeviceException /*ex*/)
            {
                //Console.WriteLine(ex.StackTrace);
            }
        }
        public static void visualise_volume()
        {
            while (true)
            {
                if (!muted)
                {
                    try
                    {
                        NAudio.CoreAudioApi.MMDeviceEnumerator MMDE = new NAudio.CoreAudioApi.MMDeviceEnumerator();
                        NAudio.CoreAudioApi.MMDevice device = MMDE.GetDefaultAudioEndpoint(NAudio.CoreAudioApi.DataFlow.Render, NAudio.CoreAudioApi.Role.Console);
                        for (int i = 0; i < inputDevice.Count; i++)
                        {
                            if (inputDevice[i].type() == Types.LAUNCHPAD_TYPE.S && inputDevice[i].IsOpen && inputDevice[i].IsReceiving)
                            {

                                double[] volume = new double[device.AudioMeterInformation.PeakValues.Count]; // so that we update all volume values at the same time

                                for (int z = 0; z < volume.Length; z++)
                                {
                                    volume[z] = device.AudioMeterInformation.PeakValues[z];
                                }

                                for (int z = 0; z < device.AudioMeterInformation.PeakValues.Count; z++) // this should just be LR?
                                {
                                    double volumePercentage = Math.Round(volume[z] * 100);

                                    if (device.AudioMeterInformation.PeakValues.Count == 2 && z == 0)
                                    { // do the left

                                        for (int q = 0; q < volumesLeft.Length - 1; q++)
                                        {
                                            volumesLeft[q] = volumesLeft[q + 1];
                                        }
                                        volumesLeft[volumesLeft.Length - 1] = volumePercentage;

                                        for (int q = 0; q < volumesLeft.Length; q++)
                                        {
                                            for (int n = 0; n < 8; n++)
                                                if (n < 8 * volumesLeft[q] / 100 && volumesLeft[q] != 0)
                                                {
                                                    if (i < outputDevice.Count && outputDevice[i].IsOpen) { outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)((q + 112) - (n * 16)), audio_colours[n]); }
                                                }
                                                else
                                                {
                                                    if (i < outputDevice.Count && outputDevice[i].IsOpen) { outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)((q + 112) - (n * 16)), 0); }
                                                }
                                        }

                                    }
                                    else if (device.AudioMeterInformation.PeakValues.Count == 2 && z == 1)
                                    { // do the right

                                        for (int q = volumesRight.Length - 1; q > 0; q--)
                                        {
                                            volumesRight[q] = volumesRight[q - 1];
                                        }
                                        volumesRight[0] = volumePercentage;

                                        for (int q = volumesRight.Length - 1; q >= 0; q--)
                                        {
                                            for (int n = 0; n < 8; n++)
                                                if (n < 8 * volumesRight[q] / 100 && volumesRight[q] != 0)
                                                {
                                                    if (i < outputDevice.Count && outputDevice[i].IsOpen) { outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)((q + 116) - (n * 16)), audio_colours[n]); }
                                                }
                                                else
                                                {
                                                    if (i < outputDevice.Count && outputDevice[i].IsOpen) { outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)((q + 116) - (n * 16)), 0); }
                                                }
                                        }

                                    }
                                    else
                                    {
                                        for (int n = 0; n < 8; n++)
                                        {
                                            if (n < 8 * volumePercentage / 100 && volumePercentage != 0)
                                            {
                                                outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)((z + 115) - (n * 16)), audio_colours[n]);
                                            }
                                            else
                                            {
                                                outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)((z + 115) - (n * 16)), 0);
                                            }

                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (DeviceException /*ex*/)
                    {
                        //Console.WriteLine(ex.StackTrace);
                    }
                }
            }
        }
        static void AudioEndpointVolume_OnVolumeNotification(NAudio.CoreAudioApi.AudioVolumeNotificationData data)
        {
            muted = data.Muted;
            if (data.Muted == true)
            {
                for (int i = 0; i < inputDevice.Count; i++)
                {
                    if (outputDevice[i].type() == Types.LAUNCHPAD_TYPE.MK2)
                    {
                        outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)39, 72);
                    }
                    else if (outputDevice[i].type() == Types.LAUNCHPAD_TYPE.S)
                    {
                        outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)88, 15);
                        for (int z = 0; z < 8; z++)
                        {
                            for (int n = 0; n < 8; n++)
                            {
                                outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)((z + 112) - (n * 16)), 0);
                            }
                        }
                    }
                }
            }
            else
            {
                for (int i = 0; i < inputDevice.Count; i++)
                {
                    if (outputDevice[i].type() == Types.LAUNCHPAD_TYPE.MK2)
                    {
                        outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)39, 0);
                    }
                    else if (outputDevice[i].type() == Types.LAUNCHPAD_TYPE.S)
                    {
                        outputDevice[i].SendNoteOn(Channel.Channel1, (Pitch)88, 0);
                    }
                }
            }
        }
        public static void _exit()
        {
            _exit(false, 0);
        }
        public static void _exit(int ExitCode)
        {
            _exit(false, ExitCode);
        }
        public static void _exit(bool kill)
        {
            _exit(kill, 0);
        }
        public static void _exit(bool kill, int ExitCode)
        {
            if (ExitCode != 0)
            {
                Environment.ExitCode = ExitCode;
            }
            if (kill == true)
            {
                Process.GetProcessById(Process.GetCurrentProcess().Id).Kill();
            }
            else
            {
                Console.WriteLine("Alright, going for a good clean shutdown.");
            }
            if (volume_visualisation.IsAlive)
            {
                Console.WriteLine("Aborting the volume visualisation.");
                volume_visualisation.Abort();
            }
            else
            {
                Console.WriteLine("Volume is not being visualised.");
            }
            try
            {
                for (int i = 0; i < inputDevice.Count; i++)
                {
                    Console.WriteLine("Cleanly shutting down device: " + inputDevice[i].Name + "/" + outputDevice[i].Name);
                    Console.WriteLine("Clearing scrolling and setting LEDs to standby.");
                    if (outputDevice[i].type() == Types.LAUNCHPAD_TYPE.MK2)
                    {
                        outputDevice[i].SendSysEx(new Byte[] { 0xF0, 0x00, 0x20, 0x29, 0x02, 0x18, 0x14, 0xF7 }); // clear scrolling for Mk2
                        outputDevice[i].SendSysEx(new Byte[] { 0xF0, 0x00, 0x20, 0x29, 0x02, 0x18, 0x0E, 0x15, 0xF7 }); // set all LEDs Mk2
                    }
                    else if (outputDevice[i].type() == Types.LAUNCHPAD_TYPE.S)
                    {
                        outputDevice[i].SendControlChange(Channel.Channel1, 0, 127); // set all LEDs S
                        outputDevice[i].SendSysEx(new Byte[] { 0xF0, 0x00, 0x20, 0x29, 0x09, 0x00, 0xF7 }); // clear scrolling for S, causes buffering issue, presumably it primes an empty scroll
                    }
                    try
                    {
                        if (outputDevice[i].IsOpen)
                        {
                            Console.WriteLine("Closing output device.");
                            outputDevice[i].Close();
                        }
                        if (inputDevice[i].IsReceiving)
                        {
                            Console.WriteLine("Stopping input device from receiving.");
                            inputDevice[i].StopReceiving();
                        }
                        if (inputDevice[i].IsOpen)
                        {
                            Console.WriteLine("Closing input device.");
                            inputDevice[i].Close();
                        }
                    }
                    catch (InvalidOperationException ex)
                    {
                        Console.WriteLine(ex.StackTrace);
                    }
                }
            }
            catch (DeviceException ex)
            {
                Console.WriteLine(ex.StackTrace);
            }
            Console.WriteLine("All done, system exiting.");
            //System.Environment.Exit(0);
        }
    }
}