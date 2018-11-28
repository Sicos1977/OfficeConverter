using System.Runtime.InteropServices;

namespace OfficeConverter.Helpers
{
    static class ProcessHelpers
    {
        /// <summary>
        /// Returns the process id for the given <paramref name="hWnd"/>
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="lpdwProcessId"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
    }
}
