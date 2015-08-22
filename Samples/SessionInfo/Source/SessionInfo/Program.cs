//-----------------------------------------------------------------------
// <copyright file="Program.cs" company="Cassia Developers">
//     Copyright © 2008 - 2015.
// </copyright>
// <version>1.01</version>
// <changes>
//     1.00 - exported from code.google.com/p/cassia
//     1.01 - updated by Hod Malkiel https://github.com/hodm/cassia
// </changes>
//-----------------------------------------------------------------------

namespace SessionInfo
{
    using System;
    using System.Collections.Generic;
    using Cassia;
    using Microsoft.Win32;
    using Res;

    /// <summary>
    /// This Class represents the Main Dos based program.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// The terminal services manager
        /// </summary>
        private static readonly ITerminalServicesManager tsManager = new TerminalServicesManager();


        /// <summary>
        /// Mains method given the command line arguments.
        /// </summary>
        /// <param name="args">The arguments.</param>
        private static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    ShowCurrentSession();
                    return;
                }
                switch (args[0].ToLower())
                {
                    case "waitforevents":
                        WaitForEvents();
                        return;
                    case "current":
                        ShowCurrentSession();
                        return;
                    case "console":
                        ShowActiveConsoleSession();
                        return;
                    case "get":
                        GetSessionInfo(args);
                        return;
                    case "listservers":
                        ListServers(args);
                        return;
                    case "listsessions":
                        ListSessions(args);
                        return;
                    case "listsessionprocesses":
                        ListSessionProcesses(args);
                        return;
                    case "listprocesses":
                        ListProcesses(args);
                        return;
                    case "killprocess":
                        KillProcess(args);
                        return;
                    case "logoff":
                        LogoffSession(args);
                        return;
                    case "disconnect":
                        DisconnectSession(args);
                        return;
                    case "message":
                        SendMessage(args);
                        return;
                    case "ask":
                        AskQuestion(args);
                        return;
                    case "shutdown":
                        Shutdown(args);
                        return;
                    case "startremotecontrol":
                        StartRemoteControl(args);
                        return;
                    case "stopremotecontrol":
                        StopRemoteControl(args);
                        return;
                    case "connect":
                        Connect(args);
                        return;
                }

                Console.WriteLine(Lang.UnknownCommand, args[0]);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        /// <summary>
        /// Connects fro one session to another session given a password to connect.
        /// SessionInfo connect [source session id] [target session id] [password]
        /// </summary>
        /// <param name="args">The arguments.</param>
        private static void Connect(string[] args)
        {
            if (args.Length < 4)
            {
                // show Usage: SessionInfo connect [source session id] [target session id] [password]
                Console.WriteLine(Lang.UsageConnect);
                return;
            }
            using (var server = tsManager.GetLocalServer())
            {
                server.Open();
                var source = server.GetSession(int.Parse(args[1]));
                var target = server.GetSession(int.Parse(args[2]));
                var password = args[3];
                source.Connect(target, password, true);
            }
        }

        private static void StopRemoteControl(string[] args)
        {
            if (args.Length < 2)
            {
                // show Usage: SessionInfo stopremotecontrol [session id]
                Console.WriteLine(Lang.UsageStopRemoteControl);
                return;
            }
            using (var server = tsManager.GetLocalServer())
            {
                server.Open();
                var session = server.GetSession(int.Parse(args[1]));
                session.StopRemoteControl();
            }
        }

        private static void StartRemoteControl(string[] args)
        {
            if (args.Length < 5)
            {
                // show Usage: SessionInfo startremotecontrol [server] [session id] [modifier] [hotkey]
                Console.WriteLine(Lang.UsageStartRemoteControl);
                return;
            }
            using (var server = GetServerFromName(args[1]))
            {
                server.Open();
                var session = server.GetSession(int.Parse(args[2]));
                var modifier =
                    (RemoteControlHotkeyModifiers)
                    Enum.Parse(typeof(RemoteControlHotkeyModifiers), args[3].Replace('+', ','), true);
                var hotkey = (ConsoleKey)Enum.Parse(typeof(ConsoleKey), args[4], true);
                session.StartRemoteControl(hotkey, modifier);
            }
        }

        private static void ShowActiveConsoleSession()
        {
            Console.WriteLine(Lang.ActiveConsoleSession);
            WriteSessionInfo(tsManager.ActiveConsoleSession);
        }

        private static void WaitForEvents()
        {
            // Show Waiting for events; press Enter to exit.
            Console.WriteLine(Lang.WaitEnterExitMsg);
            SystemEvents.SessionSwitch +=
                delegate (object sender, SessionSwitchEventArgs args) { Console.WriteLine(args.Reason); };
            Console.ReadLine();
        }

        private static void Shutdown(string[] args)
        {
            if (args.Length < 3)
            {
                // show Usage: SessionInfo shutdown [server] [shutdown type]
                Console.WriteLine(Lang.UsageShutdown);
                return;
            }
            using (var server = GetServerFromName(args[1]))
            {
                server.Open();
                var type = (ShutdownType)Enum.Parse(typeof(ShutdownType), args[2], true);
                server.Shutdown(type);
            }
        }

        private static void KillProcess(string[] args)
        {
            if (args.Length < 4)
            {
                // show Usage: SessionInfo killprocess [server] [process id] [exit code]
                Console.WriteLine(Lang.UsageKillProcess);
                return;
            }
            var processId = int.Parse(args[2]);
            var exitCode = int.Parse(args[3]);
            using (var server = GetServerFromName(args[1]))
            {
                server.Open();
                var process = server.GetProcess(processId);
                process.Kill(exitCode);
            }
        }

        private static void ListProcesses(string[] args)
        {
            if (args.Length < 2)
            {
                // show Usage: SessionInfo listprocesses [server]
                Console.WriteLine(Lang.UsageListProcesses);
                return;
            }
            using (var server = GetServerFromName(args[1]))
            {
                server.Open();
                WriteProcesses(server.GetProcesses());
            }
        }

        private static void WriteProcesses(IEnumerable<ITerminalServicesProcess> processes)
        {
            foreach (ITerminalServicesProcess process in processes)
            {
                WriteProcessInfo(process);
            }
        }

        private static void WriteProcessInfo(ITerminalServicesProcess process)
        {
            Console.WriteLine(Lang.SessionID, process.SessionId);
            Console.WriteLine(Lang.ProcessID, process.ProcessId);
            Console.WriteLine(Lang.ProcessName, process.ProcessName);
            Console.WriteLine(Lang.SID, process.SecurityIdentifier);
            Console.WriteLine(Lang.WorkingSet, process.UnderlyingProcess.WorkingSet64);
        }

        private static void ListServers(string[] args)
        {
            var domainName = (args.Length > 1 ? args[1] : null);
            foreach (ITerminalServer server in tsManager.GetServers(domainName))
            {
                Console.WriteLine(server.ServerName);
            }
        }

        private static void AskQuestion(string[] args)
        {
            if (args.Length < 8)
            {
                // show Usage: SessionInfo ask [server] [session id] [icon] [caption] [text] [timeout] [buttons]
                Console.WriteLine(Lang.UsageAsk);
                return;
            }
            var seconds = int.Parse(args[6]);
            var sessionId = int.Parse(args[2]);
            using (var server = GetServerFromName(args[1]))
            {
                server.Open();
                var session = server.GetSession(sessionId);
                var icon = (RemoteMessageBoxIcon)Enum.Parse(typeof(RemoteMessageBoxIcon), args[3], true);
                var buttons = (RemoteMessageBoxButtons)Enum.Parse(typeof(RemoteMessageBoxButtons), args[7], true);
                var result = session.MessageBox(args[5], args[4], buttons, icon, default(RemoteMessageBoxDefaultButton),
                                                default(RemoteMessageBoxOptions), TimeSpan.FromSeconds(seconds), true);
                Console.WriteLine(Lang.Response, result);
            }
        }

        private static void SendMessage(string[] args)
        {
            if (args.Length < 6)
            {
                // show Usage: SessionInfo message [server] [session id] [icon] [caption] [text]
                Console.WriteLine(Lang.UsageMessage);
                return;
            }
            var sessionId = int.Parse(args[2]);
            using (var server = GetServerFromName(args[1]))
            {
                server.Open();
                var session = server.GetSession(sessionId);
                var icon = (RemoteMessageBoxIcon)Enum.Parse(typeof(RemoteMessageBoxIcon), args[3], true);
                session.MessageBox(args[5], args[4], icon);
            }
        }

        private static void GetSessionInfo(string[] args)
        {
            if (args.Length < 3)
            {
                // show Usage: SessionInfo get [server] [session id]
                Console.WriteLine(Lang.UsageGet);
                return;
            }
            var sessionId = int.Parse(args[2]);
            using (var server = GetServerFromName(args[1]))
            {
                server.Open();
                WriteSessionInfo(server.GetSession(sessionId));
            }
        }

        private static void ListSessionProcesses(string[] args)
        {
            if (args.Length < 3)
            {
                // Usage: SessionInfo listsessionprocesses [server] [session id]
                Console.WriteLine(Lang.UsageListSessionProcesses);
                return;
            }
            var sessionId = int.Parse(args[2]);
            using (var server = GetServerFromName(args[1]))
            {
                server.Open();
                var session = server.GetSession(sessionId);
                WriteProcesses(session.GetProcesses());
            }
        }

        private static void ListSessions(string[] args)
        {
            if (args.Length < 2)
            {
                // show Usage: SessionInfo listsessions [server]
                Console.WriteLine(Lang.UsageListSessions);
                return;
            }
            using (var server = GetServerFromName(args[1]))
            {
                server.Open();
                foreach (ITerminalServicesSession session in server.GetSessions())
                {
                    WriteSessionInfo(session);
                }
            }
        }

        private static void ShowCurrentSession()
        {
            Console.WriteLine(Lang.CurrenSession);
            WriteSessionInfo(tsManager.CurrentSession);
        }

        private static void LogoffSession(string[] args)
        {
            if (args.Length < 3)
            {
                // show Usage: SessionInfo logoff [server] [session id]
                Console.WriteLine(Lang.UsageLogoff);
                return;
            }
            var serverName = args[1];
            var sessionId = int.Parse(args[2]);
            using (var server = GetServerFromName(serverName))
            {
                server.Open();
                var session = server.GetSession(sessionId);
                session.Logoff();
            }
        }

        private static ITerminalServer GetServerFromName(string serverName)
        {
            return (serverName.Equals("local", StringComparison.OrdinalIgnoreCase)
                        ? tsManager.GetLocalServer()
                        : tsManager.GetRemoteServer(serverName));
        }

        private static void DisconnectSession(string[] args)
        {
            if (args.Length < 3)
            {
                // show Usage: SessionInfo disconnect [server] [session id]
                Console.WriteLine(Lang.UsageDisconnect);
                return;
            }
            var serverName = args[1];
            var sessionId = int.Parse(args[2]);
            using (var server = GetServerFromName(serverName))
            {
                server.Open();
                var session = server.GetSession(sessionId);
                session.Disconnect();
            }
        }

        private static void WriteSessionInfo(ITerminalServicesSession session)
        {
            if (session == null)
            {
                return;
            }

            Console.WriteLine(Lang.SessionID, session.SessionId);
            if (session.UserAccount != null)
            {
                Console.WriteLine(Lang.User, session.UserAccount);
            }

            if (session.ClientIPAddress != null)
            {
                Console.WriteLine(Lang.IPAddress, session.ClientIPAddress);
            }

            if (session.RemoteEndPoint != null)
            {
                Console.WriteLine(Lang.RemoteEndpoint + session.RemoteEndPoint);
            }

            Console.WriteLine(Lang.WindowStation, session.WindowStationName);
            Console.WriteLine(Lang.ClientDirectory, session.ClientDirectory);
            Console.WriteLine(Lang.ClientBuildNumber, session.ClientBuildNumber);
            Console.WriteLine(Lang.ClientHardwareID, session.ClientHardwareId);
            Console.WriteLine(Lang.ClientProductID, session.ClientProductId);
            Console.WriteLine(Lang.ClientProtocolType, session.ClientProtocolType);
            if (session.ClientProtocolType != ClientProtocolType.Console)
            {
                // These properties often throw exceptions for the console session.
                Console.WriteLine(Lang.ApplicationName, session.ApplicationName);
                Console.WriteLine(Lang.InitialProgram, session.InitialProgram);
                Console.WriteLine(Lang.InitialWorkingDirectory, session.WorkingDirectory);
            }

            Console.WriteLine(Lang.State, session.ConnectionState);
            Console.WriteLine(Lang.ConnectTime, session.ConnectTime);
            Console.WriteLine(Lang.LogonTime, session.LoginTime);
            Console.WriteLine(Lang.LastInputTime, session.LastInputTime);
            Console.WriteLine(Lang.IdleTime, session.IdleTime);

            // show Client Display: {0}x{1} with {2} bits per pixel
            Console.WriteLine(Lang.ClientDisplay,
                            session.ClientDisplay.HorizontalResolution,
                            session.ClientDisplay.VerticalResolution,
                            session.ClientDisplay.BitsPerPixel);
            if (session.IncomingStatistics != null)
            {
                Console.Write(Lang.IncomingProtocolStatistics);
                WriteProtocolStatistics(session.IncomingStatistics);
                Console.WriteLine();
            }

            if (session.OutgoingStatistics != null)
            {
                Console.Write(Lang.OutgoingProtocolStatistics);
                WriteProtocolStatistics(session.OutgoingStatistics);
                Console.WriteLine();
            }

            Console.WriteLine();
        }

        private static void WriteProtocolStatistics(IProtocolStatistics statistics)
        {
            // show "Bytes: {0} Frames: {1} Compressed: {2}", statistics.Bytes, statistics.Frames, statistics.CompressedBytes
            Console.Write(Lang.ProtocolStatistics, statistics.Bytes, statistics.Frames, statistics.CompressedBytes);
        }
    }
}