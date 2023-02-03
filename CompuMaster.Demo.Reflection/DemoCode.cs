using CompuMaster.Reflection;
using System;

namespace CompuMaster.Demo.Reflection
{
    public class DemoCode
    {
        static void Main()
        {
            // SAMPLE: access to public property from unknown object (e.g. COM objects)
            var o = new System.Object();
            System.Console.WriteLine(
                    PublicInstanceMembers.InvokeFunction<String>(o, o.GetType(), "ToString")
                );

            // SAMPLE: access to non-public (private/internal) field from .NET object
            var r = new System.Random();
            r.Next(10, 20); //increase internal counter
            r.Next(10, 20); //increase internal counter
            System.Console.WriteLine(
                NonPublicInstanceMembers.InvokeFieldGet<Int32>(r, r.GetType(), "inext")
                );
        }
    }
}
