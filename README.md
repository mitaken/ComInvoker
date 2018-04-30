# ComInvoker

Release com object automatically

[![Build status](https://ci.appveyor.com/api/projects/status/cvt76spkwkms0ytg?svg=true)](https://ci.appveyor.com/project/mitaken/cominvoker)


## NuGet

[ComInvoker](https://www.nuget.org/packages/ComInvoker/)
```powershell
PM> Install-Package ComInvoker
```

## Usage

1. Create `ComInvoker.Invoker()` instance with __using__
1. Create COM instance using `Invoker.Invoke<T>(object com)` method

## Sample

See ComInvoker.Sample project

### Auto release

```csharp
var type = Type.GetTypeFromProgID("InternetExplorer.Application");
using (var invoker = new ComInvoker.Invoker())
{
    //Create IE Object
    var ie = invoker.Invoke<dynamic>(Activator.CreateInstance(type));
    //Show IE Window
    ie.Visible = true;
    //Open about:blank
    ie.Navigate("about:blank");
    //Close IE
    ie.Quit();
}//Release when disposing
```


### Manual release

```csharp
using (var invoker = new ComInvoker.Invoker())
{
    //Create IE Object
    var ie = invoker.Invoke<dynamic>(Activator.CreateInstance(type));
    //Show IE Window
    ie.Visible = true;
    //Open about:blank
    ie.Navigate("about:blank");
    //Close IE
    ie.Quit();
    //Release manually
    invoker.Release();
}//Release when disposing(if any com object are exists)
```
