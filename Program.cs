using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Globalization;
using System.Reflection;
using NovaAudioAssistantLib;
using NovaAudioAssistantLib.Attributes;
using NovaAudioAssistantLib.Settings;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using Timer = System.Timers.Timer;
using System.Diagnostics;

namespace NovaAudioAssistant
{
    class Program
    {
        static SpeechSynthesizer synth = new SpeechSynthesizer();
        static Dictionary<string, string> nameToUri = new Dictionary<string, string>();
        static string[] numbers = Enumerable.Range(0, 100).Select(n => n.ToString()).ToArray();
        static CommandHandler commandHandler = new CommandHandler();
        static SettingsHandler settingsHandler = new SettingsHandler("Settings.xml");
        static Thread mainThread;
        static bool consoleShown = false;
        static ContextMenu trayMenu = new ContextMenu();

        static void Main(string[] args)
        {
            // Hide Console
            var handle = GetConsoleWindow();
            DeleteMenu(GetSystemMenu(handle, false), SC_CLOSE, MF_BYCOMMAND);
            ShowWindow(handle, SW_HIDE);

            // Setup Timer for Minimize check
            Timer timer = new Timer(250);
            timer.Elapsed += (e, o) => 
            {
                if (GetMinimized(handle))
                {
                    ShowWindow(handle, SW_HIDE);
                    consoleShown = false;
                    trayMenu.MenuItems[0].Text = (consoleShown) ? "Hide Console" : "Show Console";
                }
            };
            timer.AutoReset = true;
            timer.Start();

            // Setup Tray Icon
            NotifyIcon trayIcon = new NotifyIcon();
            trayIcon.Text = "Nova Audio Assistant";
            trayIcon.Icon = new Icon(SystemIcons.Application, 40, 40);

            trayMenu.MenuItems.Add(new MenuItem("Show Console", (s, e) => 
            {
                MenuItem m = (MenuItem)s;
                consoleShown = (consoleShown) ? false : true;
                m.Text = (consoleShown) ? "Hide Console" : "Show Console";
                ShowWindow(handle, (consoleShown) ? SW_SHOW : SW_HIDE);
            }));

            // Addon Loader
            foreach (string dll in Directory.GetDirectories(@"Addons"))
            {
                string path = Directory.GetParent(dll).FullName;
                path = path + @"\" + new DirectoryInfo(dll).Name + @"\" + new DirectoryInfo(dll).Name + ".dll";
                if (!File.Exists(path)) { continue; }

                if (!File.Exists(path.Replace(".dll", ".pdb")))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("No .pdb file found for \n" + path);
                    Console.WriteLine("Exact location of error is impossible to determine");
                    Console.ForegroundColor = ConsoleColor.White;
                }
                Assembly DLL = Assembly.LoadFrom(path);
                Type[] types = DLL.GetExportedTypes();
                foreach (Type type in types)
                {
                    Addon addon = (Addon)Attribute.GetCustomAttribute(type, typeof(Addon));
                    if (addon != null)
                    {
                        Console.ForegroundColor = ConsoleColor.Blue;
                        Console.WriteLine($"{addon.Name} By {addon.Creator}");
                        Console.WriteLine(addon.Desc);
                        Console.WriteLine($"Version: {addon.Version}");
                        Console.ForegroundColor = ConsoleColor.White;
                        dynamic c = Activator.CreateInstance(type);
                        try
                        {
                            c.Init(commandHandler, synth, settingsHandler);
                            try
                            {
                                c.InitUI(trayMenu);
                            }
                            catch (RuntimeBinderException) { }
                            catch (Exception e)
                            {
                                Console.ForegroundColor = ConsoleColor.DarkRed;
                                Console.WriteLine("Failed to load " + path);
                                Console.WriteLine("(" + e.Message + ") thrown in " + e.TargetSite.Name + " at line " + GetLineNumber(e));
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("Sucessfully loaded!");
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                        catch (Exception e)
                        {
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine("Failed to load " + path);
                            Console.WriteLine("Method Init Missing or Incorrect or other error has occured");
                            Console.WriteLine("(" + e.Message + ") thrown in " + e.TargetSite.Name + " at line " + GetLineNumber(e));
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.DarkMagenta;
                        Console.WriteLine("Failed to load " + path);
                        Console.WriteLine("File doesnt contain Addon Attribute/Code");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                }
            }

            // Setup Listening Thread
            mainThread = new Thread(() =>
            {
                synth.SetOutputToDefaultAudioDevice();

                //Main Listening loop
                using (SpeechRecognitionEngine engine = new SpeechRecognitionEngine(new CultureInfo("en-US")))
                {
                    Choices pre = commandHandler.GetPreCommands();
                    Choices command = commandHandler.GetCommands();
                    Choices sub = commandHandler.GetSubCommands();
                    sub.Add(" ");
                    GrammarBuilder bg = new GrammarBuilder();
                    bg.Append(pre);
                    bg.Append(command);
                    bg.Append(sub);
                    engine.LoadGrammar(new Grammar(bg));
                    engine.SpeechRecognized += Engine_SpeechRecognized;
                    engine.SetInputToDefaultAudioDevice();
                    engine.RecognizeAsync(RecognizeMode.Multiple);
                    while (true)
                    {
                        Console.ReadLine();
                    }
                }
            });
            mainThread.Start();

            // Finish Tray Icon Setup
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;

            Console.Title = "Nova Audio Assistant";

            trayMenu.MenuItems.Add(new MenuItem("Quit", Quit));

            Application.Run();
        }

        private static void Quit(object sender, EventArgs e)
        {
            settingsHandler.SaveSettings();
            mainThread.Abort();
            Application.Exit();
            Environment.Exit(0);
        }

        static int GetLineNumber(Exception e)
        {
            var st = new StackTrace(e, true);
            var frame = st.GetFrame(0);
            return frame.GetFileLineNumber();
        }

        static bool waiting = false;
        private static void Engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (waiting)
            {
                return;
            }
            if (e.Result.Confidence >= commandHandler.GetConfidence())
            {
                Console.WriteLine(e.Result.Text);
                string[] command = e.Result.Text.Split(' ');
                waiting = true;
                commandHandler.RunCommand(command[0], command[1], (command.Length > 2)? command[2] : "");
                waiting = false;
            }
            else
            {
                Console.WriteLine("Im not that confinent");
            }
        }

        #region Window Management
        private const int MF_BYCOMMAND = 0x00000000;
        public const int SC_CLOSE = 0xF060;

        [DllImport("user32.dll")]
        public static extern int DeleteMenu(IntPtr hMenu, int nPosition, int wFlags);

        [DllImport("user32.dll")]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

        private struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public Point ptMinPosition;
            public Point ptMaxPosition;
            public Rectangle rcNormalPosition;
        }

        public static bool GetMinimized(IntPtr handle)
        {
            WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf(placement);
            GetWindowPlacement(handle, ref placement);
            return placement.showCmd == SW_SHOWMINIMIZED;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const UInt32 SW_SHOWNORMAL = 1;
        const UInt32 SW_NORMAL = 1;
        const UInt32 SW_SHOWMINIMIZED = 2;
        const UInt32 SW_SHOWMAXIMIZED = 3;
        const UInt32 SW_MAXIMIZE = 3;
        const UInt32 SW_SHOWNOACTIVATE = 4;
        const int SW_SHOW = 5;
        const UInt32 SW_MINIMIZE = 6;
        const UInt32 SW_SHOWMINNOACTIVE = 7;
        const UInt32 SW_SHOWNA = 8;
        const UInt32 SW_RESTORE = 9;
        #endregion
    }
}
