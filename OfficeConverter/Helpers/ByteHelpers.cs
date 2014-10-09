namespace OfficeConverter.Helpers
{
    /// <summary>
    /// This class contains byte helper methods that are not available in the .NET framework
    /// </summary>
    internal static class ByteHelpers
    {
        /// <summary>
        /// Returns true if the given <paramref name="b"/> is set on position <paramref name="position"/>
        /// </summary>
        /// <param name="b"></param>
        /// <param name="position">The zero based position to check</param>
        /// <returns></returns>
        public static bool IsBitSet(this byte b, int position)
        {
            return (b & (1 << position)) != 0;
        }
    }
}
