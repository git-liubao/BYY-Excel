using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PackageExcel
{
    class MyComparer : IComparer
    {

        // Calls CaseInsensitiveComparer.Compare with the parameters reversed.
        int IComparer.Compare(Object x, Object y)
        {
            string sx = x as string;
            string sy = y as string;
            sx = sx.Substring(sx.LastIndexOf('\\') + 1);
            sy = sy.Substring(sy.LastIndexOf('\\') + 1);

            var idx = Int32.Parse(sx.Split("_".ToArray())[0]);
            var idy = Int32.Parse(sy.Split("_".ToArray())[0]);
            return idx - idy;
        }

    }
}
