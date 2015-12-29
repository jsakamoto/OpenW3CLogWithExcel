using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenW3CLogWithExcel
{
    internal static class Extensions
    {
        public static void Each<T>(this IEnumerable<T> collection, Action<T, int> action)
        {
            var index = 0;
            foreach (var item in collection)
            {
                action(item, index);
                index++;
            }
        }
    }
}
