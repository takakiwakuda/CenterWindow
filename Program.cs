// (c) Takaki Wakuda.

using System.CommandLine;
using System.CommandLine.Builder;
using System.CommandLine.Hosting;
using System.CommandLine.Invocation;
using System.CommandLine.IO;
using System.CommandLine.Parsing;
using System.Runtime.InteropServices;
using System.Text;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Graphics.Gdi;
using Windows.Win32.UI.HiDpi;

namespace CenterWindow;

static class Program
{
    static async Task Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        await new CommandLineBuilder(new CenterWindowCommand()).UseDefaults()
            .UseHost(host => host.UseCommandHandler<CenterWindowCommand, CenterWindowCommandHandler>())
            .Build()
            .InvokeAsync(args);
    }
}

sealed class CenterWindowCommand : RootCommand
{
    public CenterWindowCommand() : base("Center the window.")
    {
        AddArgument(new Argument<int>("handle", "Specify a window handle"));
        AddOption(new Option<bool>("--disable-dpi-awareness"));
        AddOption(new Option<bool>("--dry-run"));
        AddOption(new Option<bool>("--use-work-area"));
        AddOption(new Option<bool>(["-v", "--verbose"]));
    }
}

sealed class CenterWindowCommandHandler : ICommandHandler
{
    private static class ExitStatus
    {
        internal static Task<int> Success => Task.FromResult(0);
        internal static Task<int> Failure => Task.FromResult(1);
    }

    private static DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 => new(-4);

    public int Handle { get; set; }
    public bool DisableDpiAwareness { get; set; }
    public bool DryRun { get; set; }
    public bool UseWorkArea { get; set; }
    public bool Verbose { get; set; }

    public int Invoke(InvocationContext context)
    {
        throw new NotImplementedException();
    }

    public Task<int> InvokeAsync(InvocationContext context)
    {
        var stdout = context.Console.Out;
        var stderr = context.Console.Error;

        if (DisableDpiAwareness)
        {
            stdout.WriteLine("DPI awareness context disabled.");
        }
        else
        {
            if (!PInvoke.SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2))
            {
                stderr.WritePInvokeErrorMessage();
                return ExitStatus.Failure;
            }
            stdout.WriteLine("DPI awareness context enabled.");
        }

        var hWnd = new HWND(Handle);
        if (!PInvoke.GetWindowRect(hWnd, out RECT rect))
        {
            stderr.WritePInvokeErrorMessage();
            return ExitStatus.Failure;
        }

        var monitor = PInvoke.MonitorFromWindow(hWnd, MONITOR_FROM_FLAGS.MONITOR_DEFAULTTONEAREST);
        var monitorInfo = new MONITORINFO
        {
            cbSize = (uint)Marshal.SizeOf<MONITORINFO>()
        };
        if (!PInvoke.GetMonitorInfo(monitor, ref monitorInfo))
        {
            stderr.WritePInvokeErrorMessage();
            return ExitStatus.Failure;
        }

        int screenWidth;
        int screenHeight;
        if (UseWorkArea)
        {
            screenWidth = monitorInfo.rcWork.Width;
            screenHeight = monitorInfo.rcWork.Height;
        }
        else
        {
            screenWidth = monitorInfo.rcMonitor.Width;
            screenHeight = monitorInfo.rcMonitor.Height;
        }

        int x = (screenWidth - rect.Width) / 2;
        int y = (screenHeight - rect.Height) / 2;

        if (Verbose)
        {
            stdout.WriteLine();
            stdout.WriteLine("Window Information:");
            stdout.WriteLine($"  Title     : {GetWindowTitle(hWnd)}");
            stdout.WriteLine($"  Handle    : {Handle}");
            stdout.WriteLine($"  Top       : {rect.top}");
            stdout.WriteLine($"  Bottom    : {rect.bottom}");
            stdout.WriteLine($"  Right     : {rect.right}");
            stdout.WriteLine($"  Left      : {rect.left}");
            stdout.WriteLine($"  Width     : {rect.Width}");
            stdout.WriteLine($"  Height    : {rect.Height}");
            stdout.WriteLine();
            stdout.WriteLine("Screen Information:");
            stdout.WriteLine($"  Width     : {screenWidth}");
            stdout.WriteLine($"  Height    : {screenHeight}");
            stdout.WriteLine($"  Scale (%) : {PInvoke.GetDpiForWindow(hWnd) / 96.0 * 100}");
            stdout.WriteLine();
            stdout.WriteLine("New Position:");
            stdout.WriteLine($"  X         : {x}");
            stdout.WriteLine($"  Y         : {y}");
            stdout.WriteLine();
        }

        if (DryRun)
        {
            stdout.WriteLine($"Move the window ({Handle}) to (X={x}, Y={y}), but no changes were applied.");
        }
        else
        {
            if (!PInvoke.MoveWindow(hWnd, x, y, rect.Width, rect.Height, true))
            {
                stderr.WritePInvokeErrorMessage();
                return ExitStatus.Failure;
            }
            stdout.WriteLine($"Successfully moved the window ({Handle}) to (X={x}, Y={y}).");
        }

        return ExitStatus.Success;
    }

    private static unsafe string GetWindowTitle(HWND hWnd)
    {
        int length = PInvoke.GetWindowTextLength(hWnd);
        if (length == 0)
        {
            return string.Empty;
        }

        Span<char> buffer = stackalloc char[length + 1];
        int copied;

        fixed (char* pBuffer = buffer)
        {
            copied = PInvoke.GetWindowText(hWnd, pBuffer, buffer.Length);
        }

        if (copied == 0)
        {
            return string.Empty;
        }

        return new string(buffer[..copied]);
    }
}

file static class IStandardStreamWriterExtensions
{
    internal static void WritePInvokeErrorMessage(this IStandardStreamWriter writer)
    {
        int error = Marshal.GetLastPInvokeError();
        string message = Marshal.GetPInvokeErrorMessage(error);
        var originalColor = Console.ForegroundColor;

        try
        {
            Console.ForegroundColor = ConsoleColor.Red;
            writer.WriteLine($"{message} (Error={error})");
        }
        finally
        {
            Console.ForegroundColor = originalColor;
        }
    }
}
