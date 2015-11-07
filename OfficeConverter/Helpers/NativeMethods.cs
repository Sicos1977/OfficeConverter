using System.Runtime.InteropServices;

/*
   Copyright 2014-2015 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

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
