using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using HWND = System.IntPtr;

namespace OfficeConverter.Helpers
{
    static class ProcessHelpers
    {
        public delegate bool EnumedWindow(IntPtr handleWindow, ArrayList handles);

        /// <summary>
        /// Returns the process id for the given <paramref name="hWnd"/>
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="lpdwProcessId"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        private delegate bool EnumWindowsProc(HWND hWnd, int lParam);

        [DllImport("USER32.DLL")]
        private static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(HWND hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowTextLength(HWND hWnd);

        [DllImport("USER32.DLL")]
        private static extern bool IsWindowVisible(HWND hWnd);

        [DllImport("USER32.DLL")]
        private static extern IntPtr GetShellWindow();

        /// <summary>
        /// Returns the process id for the Windows with the given <paramref name="title"/>
        /// </summary>
        /// <param name="title"></param>
        /// <returns></returns>
        public static int? GetProcessIdByWindowTitle(string title)
        {
            var shellWindow = GetShellWindow();
            var windows = new Dictionary<HWND, string>();

            EnumWindows(delegate(HWND hWnd, int lParam)
            {
                if (hWnd == shellWindow) return true;

                var length = GetWindowTextLength(hWnd);

                var stringBuilder = new StringBuilder(length);
                GetWindowText(hWnd, stringBuilder, length + 1);

                windows[hWnd] = stringBuilder.ToString();
                return true;

            }, 0);

            if (!windows.ContainsValue(title)) return null;
            var window = windows.First(m => m.Value == title);
            GetWindowThreadProcessId(window.Key.ToInt32(), out var processId);
            return processId;
        }
    }
}
