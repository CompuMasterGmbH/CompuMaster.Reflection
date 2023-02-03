# Reflection
Reflection simplified for all-days-purposes

## When you need this COM interop library
* Access to public members of COM objects when not type libraries ("interop assemblies") are present and late-binding doesn't work (e.g. a late-binding-call to Excel.Quit failed)
* Access to private/non-public members of .NET assemblies

## Sample implementation
```C#
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
```