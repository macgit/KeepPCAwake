Add-Type @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PInvoke.Win32 {

    public static class UserInput {

        [DllImport("user32.dll", SetLastError=false)]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO {
            public uint cbSize;
            public int dwTime;
        }

        public static DateTime LastInput {
            get {
                DateTime bootTime = DateTime.UtcNow.AddMilliseconds(-Environment.TickCount);
                DateTime lastInput = bootTime.AddMilliseconds(LastInputTicks);
                return lastInput;
            }
        }

        public static TimeSpan IdleTime {
            get {
                return DateTime.UtcNow.Subtract(LastInput);
            }
        }

        public static int LastInputTicks {
            get {
                LASTINPUTINFO lii = new LASTINPUTINFO();
                lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
                GetLastInputInfo(ref lii);
                return lii.dwTime;
            }
        }
    }
}
'@

$myshell = New-Object -com "Wscript.Shell"
$null | Set-Content "C:\Users\sakmishra\Documents\mouse_test.log"

while($true){

    $IdleTime=[PInvoke.Win32.UserInput]::IdleTime
    $LastInput=[PInvoke.Win32.UserInput]::LastInput
    $DateTime=Get-Date -Format "dd-MM-yyyy HH:mm:ss"
 
    #Write-Host ("Last input " + [PInvoke.Win32.UserInput]::LastInput)
    #Write-Host ("Idle for " + [PInvoke.Win32.UserInput]::IdleTime)

    "LoopEnteredInfo - DateTime($DateTime) : IdleTime($IdleTime) : LastInput($LastInput)" | Add-Content "C:\Users\sakmishra\Documents\mouse_test.log"

    $Minutes = ([TimeSpan]::Parse([PInvoke.Win32.UserInput]::IdleTime)).Minutes

    if($Minutes -ge 4){
        1..2 | ForEach-Object{
        $myshell.sendkeys("^{ESC}");
        Start-Sleep 1
        }
        "ResetInfo - DateTime($DateTime) : IdleTime($IdleTime) : LastInput($LastInput)" | Add-Content "C:\Users\sakmishra\Documents\mouse_test.log"
    }
    Start-Sleep 60
}
