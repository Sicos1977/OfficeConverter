using System.Runtime.InteropServices;

namespace OfficeConverter.Helpers
{
    internal static class NativeMethods
    {
        /// <summary>
        /// Returns true when running on a server system
        /// </summary>
        /// <returns></returns>
        public static bool IsWindowsServer()
        {
            return IsOS(OsAnyserver);
        }

        const int OsAnyserver = 29;

        [DllImport("shlwapi.dll", SetLastError = true, EntryPoint = "#437")]
        private static extern bool IsOS(int os);
    }
}
