using System.Runtime.InteropServices;

namespace OfficeConverter.Helpers
{
    static class ProcessHelpers
    {
        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
    }
}
