
using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace TestRTD;

[ComVisible(true)]
[Guid("BA9AC84B-C7FC-41CF-8B2F-1764EB773D4B")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IServer : IRtdServer
{

}