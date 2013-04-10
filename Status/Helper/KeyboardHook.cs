/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 */
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace VmcController.MceState {
	public class KeyboardHook : IDisposable {
		#region Member Variables

		private const int WH_KEYBOARD_LL = 13;
		private const int WM_KEYDOWN = 0x0100;
		private IntPtr m_hookID = IntPtr.Zero;
		private LowLevelKeyboardProc m_callbackDelegate;
		private bool m_disposed;

		#endregion

		#region Events

		/// <summary>
		/// Keyboard key press event handler
		/// </summary>
		public event EventHandler<KeyboardHookEventArgs> KeyPress;

		#endregion

		#region Constructor

		/// <summary>
		/// Initializes a new instance of the <see cref="KeyboardHook"/> class.
		/// </summary>
		public KeyboardHook() {
			if (m_hookID == IntPtr.Zero) {
				using (Process process = Process.GetCurrentProcess())
				using (ProcessModule module = process.MainModule) {
					IntPtr hModule = NativeMethods.GetModuleHandle(module.ModuleName);
					m_callbackDelegate = HookCallback;
					m_hookID = NativeMethods.SetWindowsHookEx(WH_KEYBOARD_LL, m_callbackDelegate, hModule, 0);
				}
			}
		}

		#endregion

		#region Public Methods

		/// <summary>
		/// Determines whether the keyboard is hooked.
		/// </summary>
		/// <returns>
		/// 	<c>true</c> if this instance is hooked; otherwise, <c>false</c>.
		/// </returns>
		public bool IsHooked() {
			return (m_hookID != IntPtr.Zero && m_callbackDelegate != null);
		}

		#endregion

		#region Private Methods

		/// <summary>
		/// Keyboard hook callback.
		/// </summary>
		/// <param name="nCode">The n code.</param>
		/// <param name="wParam">The w param.</param>
		/// <param name="lParam">The l param.</param>
		/// <returns></returns>
		private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam) {
			if (KeyPress != null && nCode >= 0 && wParam == (IntPtr) WM_KEYDOWN) {
				IntPtr ehshellWnd = GetWindowHandle();
				if ((ehshellWnd != IntPtr.Zero) && (ehshellWnd == NativeMethods.GetForegroundWindow())) {
					int vkCode = Marshal.ReadInt32(lParam);
					KeyPress(this, new KeyboardHookEventArgs(vkCode));
				}
			}
			return NativeMethods.CallNextHookEx(m_hookID, nCode, wParam, lParam);
		}

		/// <summary>
		/// Helper method to get the window handle for the main MCE shell in our session
		/// </summary>
		/// <returns></returns>
		private IntPtr GetWindowHandle() {
			IntPtr hwnd = IntPtr.Zero;
			int mySession;

			using (Process currentProcess = Process.GetCurrentProcess())
				mySession = currentProcess.SessionId;

			foreach (Process p in Process.GetProcessesByName("ehshell")) {
				if (p.SessionId == mySession) {
					hwnd = p.MainWindowHandle;
				}
				p.Dispose();
			}
			return hwnd;
		}

		#endregion

		#region IDisposable Members

		/// <summary>
		/// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
		/// </summary>
		public void Dispose() {
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		/// <summary>
		/// Dispose(bool disposing) executes in two distinct scenarios.
		/// If disposing equals true, the method has been called directly
		/// or indirectly by a user's code. Managed and unmanaged resources
		/// can be disposed.
		/// If disposing equals false, the method has been called by the 
		/// runtime from inside the finalizer and you should not reference 
		/// other objects. Only unmanaged resources can be disposed.
		/// </summary>
		/// <param name="disposing">if set to <c>true</c> [disposing].</param>
		protected virtual void Dispose(bool disposing) {
			// Check to see if Dispose has already been called.
			if (!this.m_disposed) {
				if (m_hookID != IntPtr.Zero) {
					NativeMethods.UnhookWindowsHookEx(m_hookID);
					m_hookID = IntPtr.Zero;
				}
			}
			m_disposed = true;
		}

		/// <summary>
		/// Releases unmanaged resources and performs other cleanup operations before the
		/// <see cref="KeyboardHook"/> is reclaimed by garbage collection.
		/// </summary>
		~KeyboardHook() {
			Dispose(false);
		}

		#endregion

		#region Win32 Imports

		internal delegate IntPtr LowLevelKeyboardProc(int code, IntPtr wParam, IntPtr lParam);

		internal static class NativeMethods {
			[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
			internal static extern IntPtr GetForegroundWindow();

			[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
			internal static extern IntPtr SetWindowsHookEx(int idHook,
			                                               LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

			[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
			[return: MarshalAs(UnmanagedType.Bool)]
			internal static extern bool UnhookWindowsHookEx(IntPtr hhk);

			[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
			internal static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
			                                             IntPtr wParam, IntPtr lParam);

			[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
			internal static extern IntPtr GetModuleHandle(string lpModuleName);
		}

		#endregion
	}

	#region KeyboardHookEventArgs

	public class KeyboardHookEventArgs : EventArgs {
		private int m_vkCode;

		/// <summary>
		/// Initializes a new instance of the <see cref="KeyboardHookEventArgs"/> class.
		/// </summary>
		/// <param name="vkCode">The vk code.</param>
		public KeyboardHookEventArgs(int vkCode) {
			m_vkCode = vkCode;
		}

		/// <summary>
		/// Gets the vk code.
		/// </summary>
		/// <value>The vk code.</value>
		public int vkCode {
			get { return m_vkCode; }
		}
	}

	#endregion
}