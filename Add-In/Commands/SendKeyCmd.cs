/*
 * Copyright (c) 2007 Jonathan Bradshaw / 2012 Gert-Jan Niewenhuijse
 * 
 * This software is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely, subject to the following restrictions:
 * 
 * 1. The origin of this software must not be misrepresented; you must not claim that you wrote 
 *    the original software. If you use this software in a product, an acknowledgment in the 
 *    product documentation would be appreciated but is not required.
 * 2. Altered source versions must be plainly marked as such, and must not be misrepresented as
 *    being the original software.
 * 3. This notice may not be removed or altered from any source distribution.
 * 
 */
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.VisualBasic;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for MsgBox command.
	/// </summary>
	public class SendKeyCmd : ICommand
    {
        private ushort vk = 0;
		public bool			Shift = false;
		public bool			Ctrl = false;
		public bool			Alt = false;
		public bool			Win = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="SendKey"/> class.
        /// </summary>
        /// <param name="vk">The vk.</param>
        public SendKeyCmd(ushort vk)
        {
            this.vk = vk;
            this.Shift = false;
            this.Ctrl = false;
            this.Alt = false;
            this.Win = false;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SendKey"/> class.
        /// </summary>
        /// <param name="vk">The vk.</param>
        /// <param name="Shift">if set to <c>true</c> [shift].</param>
        /// <param name="Ctrl">if set to <c>true</c> [CTRL].</param>
        /// <param name="Alt">if set to <c>true</c> [alt].</param>
        public SendKeyCmd(ushort vk, bool Shift, bool Ctrl, bool Alt)
		{
			this.vk = vk;
			this.Shift = Shift;
			this.Ctrl = Ctrl;
			this.Alt = Alt;
			this.Win = false;
		}

        /// <summary>
        /// Initializes a new instance of the <see cref="SendKey"/> class.
        /// </summary>
        /// <param name="vk">The vk.</param>
        /// <param name="Shift">if set to <c>true</c> [shift].</param>
        /// <param name="Ctrl">if set to <c>true</c> [CTRL].</param>
        /// <param name="Alt">if set to <c>true</c> [alt].</param>
        /// <param name="Win">if set to <c>true</c> [win].</param>
        public SendKeyCmd(ushort vk, bool Shift, bool Ctrl, bool Alt, bool Win) 
		{
			this.vk = vk;
			this.Shift = Shift;
			this.Ctrl = Ctrl;
			this.Alt = Alt;
			this.Win = Win;
		}

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "- sends key to application";
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            OpResult opResult = new OpResult();

            try
            {
                if (!GetWindowToForeground())
                {
                    opResult.StatusCode = OpStatusCode.BadRequest;
                    return opResult;
                }

                if (Shift)
                {
                    NativeMethods.SendKeyDown(0x10);
                }
                if (Ctrl)
                {
                    NativeMethods.SendKeyDown(0x11);
                }
                if (Alt)
                {
                    NativeMethods.SendKeyDown(0x12);
                }
                if (Win)
                {
                    NativeMethods.SendKeyDown(0x5b);
                }

                NativeMethods.SendKeyStroke(vk);

                if (Shift)
                {
                    NativeMethods.SendKeyUp(0x10);
                }
                if (Ctrl)
                {
                    NativeMethods.SendKeyUp(0x11);
                }
                if (Alt)
                {
                    NativeMethods.SendKeyUp(0x12);
                }
                if (Win)
                {
                    NativeMethods.SendKeyUp(0x5b);
                }

                opResult.StatusCode = OpStatusCode.Success;
            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = ex.Message;
            }
            return opResult;
        }

        #endregion

        /// <summary>
        /// Helper method to get the window handle for the main MCE shell in our session
        /// </summary>
        /// <returns></returns>
        private static bool GetWindowToForeground()
        {
            int PID;

            PID = GetPID();
            if (PID == 0)
            {
                    return false;
            }

            Interaction.AppActivate(PID);
            return true;
        }

        /// <summary>
        /// Helper method to get the window handle for the main MCE shell in our session
        /// </summary>
        /// <returns></returns>
        private static int GetPID()
        {
            int PID = 0;
            int mySession;

            using (Process currentProcess = Process.GetCurrentProcess())
                mySession = currentProcess.SessionId;

            foreach (Process p in Process.GetProcessesByName("ehshell"))
            {
                if (p.SessionId == mySession)
                {
                    if (NativeMethods.IsIconic(p.MainWindowHandle))
                    {
                        NativeMethods.ShowWindow(p.MainWindowHandle, NativeMethods.SW_RESTORE);
                    }
                    else
                    {
                        PID = p.Id;
                    }
                }
                p.Dispose();
            }
            return PID;
        }

        #region Win32 Helpers
        static class NativeMethods
        {
            private const int INPUT_KEYBOARD = 1;
            private const uint KEY_UP = 0x0002;

            [StructLayout(LayoutKind.Sequential)]
            private struct KEYBDINPUT
            {
                public ushort wVk;
                public ushort wScan;
                public uint dwFlags;
                public uint time;
                public IntPtr dwExtraInfo;
            } ;

            [StructLayout(LayoutKind.Sequential)]
            private struct MOUSEINPUT
            {
                public int dx;
                public int dy;
                public uint mouseData;
                public uint dwFlags;
                public uint time;
                public IntPtr dwExtraInfo;
            };

            [StructLayout(LayoutKind.Sequential)]
            struct HARDWAREINPUT
            {
                uint uMsg;
                ushort wParamL;
                ushort wParamH;
            }

            [StructLayout(LayoutKind.Explicit, Size = 28)]
            private struct INPUT32
            {
                [FieldOffset(0)]
                public int type;
                [FieldOffset(4)]
                MOUSEINPUT mi;
                [FieldOffset(4)]
                public KEYBDINPUT ki;
                [FieldOffset(4)]
                HARDWAREINPUT hi;
            };           

            [StructLayout(LayoutKind.Explicit)]
            private struct INPUT64
            {
                [FieldOffset(0)]
                public uint type;
                [FieldOffset(8)]
                public MOUSEINPUT mi;
                [FieldOffset(8)]
                public KEYBDINPUT ki;
                [FieldOffset(8)]
                HARDWAREINPUT hi;
            };

            public static void SendKeyStroke(ushort key)
            {
                SendKeyDown(key);
                SendKeyUp(key);
            }

            public static void SendKeyDown(ushort key)
            {
                if (IntPtr.Size == 8)
                {
                    INPUT64 input64 = new INPUT64();
                    input64.type = INPUT_KEYBOARD;
                    input64.ki.wVk = key;
                    input64.ki.wScan = 0;
                    input64.ki.time = 0;
                    input64.ki.dwFlags = 0;
                    input64.ki.dwExtraInfo = IntPtr.Zero;

                    // Key down the actual key-code
                    SendInput64.SendInput(1, ref input64, Marshal.SizeOf(input64));
                }
                else
                {
                    INPUT32 input32 = new INPUT32();
                    input32.type = INPUT_KEYBOARD;
                    input32.ki.wVk = key;
                    input32.ki.wScan = 0;
                    input32.ki.time = 0;
                    input32.ki.dwFlags = 0;
                    input32.ki.dwExtraInfo = GetMessageExtraInfo();

                    // Key down the actual key-code
                    SendInput32.SendInput(1, ref input32, Marshal.SizeOf(input32));
                }
            }

            public static void SendKeyUp(ushort key)
            {
                if (IntPtr.Size == 8)
                {
                    INPUT64 input64 = new INPUT64();
                    input64.type = INPUT_KEYBOARD;
                    input64.ki.wVk = key;
                    input64.ki.wScan = 0;
                    input64.ki.time = 0;
                    input64.ki.dwFlags = KEY_UP;
                    input64.ki.dwExtraInfo = IntPtr.Zero;

                    // Key up the actual key-code
                    SendInput64.SendInput(1, ref input64, Marshal.SizeOf(input64));
                }
                else
                {
                    INPUT32 input32 = new INPUT32();
                    input32.type = INPUT_KEYBOARD;
                    input32.ki.wVk = key;
                    input32.ki.wScan = 0;
                    input32.ki.time = 0;
                    input32.ki.dwFlags = KEY_UP;
                    input32.ki.dwExtraInfo = GetMessageExtraInfo();

                    // Key up the actual key-code
                    SendInput32.SendInput(1, ref input32, Marshal.SizeOf(input32));
                }
            }

            private class SendInput32
            {
                [DllImport("user32.dll")]
                public static extern uint SendInput(uint nInputs, ref INPUT32 pInputs, int cbSize);
            }

            private class SendInput64
            {
                [DllImport("user32.dll")]
                public static extern uint SendInput(uint nInputs, ref INPUT64 pInputs, int cbSize);
            }

            [DllImport("user32.dll")]
            internal static extern IntPtr GetMessageExtraInfo();

            // Activate an application window.
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            // Show a window in a given visibility (maximize, minimize, restore etc.)
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            public const int SW_MAXIMIZE = 3;
            public const int SW_RESTORE = 9;

            // Check if a window is minimized
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool IsIconic(IntPtr hWnd);

        }
        #endregion
    }
}
