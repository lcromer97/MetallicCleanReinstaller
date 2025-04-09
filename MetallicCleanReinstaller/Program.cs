using System.Diagnostics;
using System.Security.Principal;
using System.ServiceProcess;
using Microsoft.Win32;

namespace MintyLabs.MetallicCleanReinstaller;

public class Program {
    internal static readonly string[] servicesToBeRemoved = [
        "GxBlr(Instance001)", 
        "GxClMgrS(Instance001)",
        "GxCVD(Instance001)",
        "GxFWD(Instance001)",
        "GxVssHWProv(Instance001)",
        "GxVssProv(Instance001)" ];

    internal static readonly string[] foldersToBeDeleted = [ 
        @"C:\Program Files\CommVault",
        @"C:\Program Files\Metallic",
        @"C:\Program Files (x86)\CommVault",
        @"C:\Program Files (x86)\Metallic",
        @"C:\ProgramData\Commvault Systems" ];

    internal static readonly string[] registryKeysToBeRemoved = [
        @"SOFTWARE\CommVault Systems",
        @"SOFTWARE\GalaxyInstallerFlags",
        @"SOFTWARE\GalaxyRemoteInstall" ];

    internal const string HostInstaller = @"<redacted>";
    internal const string VirtualInstaller = @"<redacted>";

    internal static List<string> foundServicesToCheckLater = [];
    internal static FileInfo? windows_Server64;

    public static async Task Main(string[]? args = null) {
        // Check if program was ran as Admin. If false, relaunch with Admin UAC prompt
        if (!IsAdministrator()) {
            Console.WriteLine("Restarting with administrative privileges...");
            RestartAsAdmin();
            return;
        }

        Log("Running as administrator.");
        
        // Delete Folders
        await DeleteFolders(foldersToBeDeleted);

        // Delete Registry Keys
        await DeleteRegistryKeys(registryKeysToBeRemoved);

        // Delete Metallic Services
        await DeleteServices(servicesToBeRemoved);

        // Prompt User to specify if current machine is a Host or Virtual Machine
        Console.Write("Is this a Host or Virtual Server? (H/V): ");
        var machineTypeChoice = Console.ReadLine()?.ToUpper();
        var fileToCopy = machineTypeChoice == "V" ? VirtualInstaller : HostInstaller;
        var destination = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileName(fileToCopy));
        // Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\Temp\" + Path.GetFileName(fileToCopy);

        // Copy Commvault installer to machine
        CopyFile(fileToCopy, destination);

        // Tell user that they will need this file as a part of the installation, then open notepad with that file
        Log("Opening the Commvault Authorization Key.txt file, you will need this later...");
        OpenNotepadForKey(@"<redacted>");

        // Execute Commvault installer
        await ExecuteFile(destination);

        // Wait for execution to fully exit
        await Task.Delay(TimeSpan.FromSeconds(5));

        // Check and compare the old list of services to the new current set
        CheckNewlyInstalledServices();

        // remove the large file
        CleanupFiles();

        Console.Write("Would you like to reboot this server? (Y/N): ");
        var rebootChoice = Console.ReadLine()?.ToUpper();

        if (rebootChoice == "Y") {
            Console.Write("Type 'REBOOT' to restart the machine: ");
            if (Console.ReadLine() == "REBOOT") {
                Log("Rebooting system...");
                Process.Start("shutdown", "/f /r /t 0");
            }
        }
    }

    /// <summary>
    /// Starts current program as admin
    /// </summary>
    private static void RestartAsAdmin() {
        var exePath = Environment.ProcessPath;
        Process.Start(new ProcessStartInfo() {
            FileName = exePath,
            UseShellExecute = true,
            Verb = "runas"
        });
    }

    /// <summary>
    /// Deletes a list of folders
    /// </summary>
    private static async Task DeleteFolders(string[] folders) {
        foreach (var folder in folders) {
            var directory = new DirectoryInfo(folder);

            if (directory.Exists) {
                Log($"Deleting \"{folder}\"...");
                Directory.Delete(directory.FullName, true);
                Log($"Deleted folder: \"{folder}\"", LogLevel.Success);
                await Task.Delay(TimeSpan.FromSeconds(5));
                continue;
            }
            else {
                Log($"Could not find folder \"{directory.FullName}\"", LogLevel.Warning);
                continue;
            }
        }
    }

    /// <summary>
    /// Deletes a list of Registry Keys
    /// </summary>
    private static async Task DeleteRegistryKeys(string[] keys) {
        foreach (var key in keys) {
            try {
                using var regKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(key, true);
                if (regKey is not null) {
                    DeleteSubKeys(regKey);
                    RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).DeleteSubKeyTree(key);
                    Log($"Deleted registry key: \"HKEY_LOCAL_MACHINE\\{key}\"", LogLevel.Success);
                    await Task.Delay(TimeSpan.FromSeconds(1));
                }
                else {
                    Log($"Registry Key: \"HKEY_LOCAL_MACHINE\\{key}\", could not be found", LogLevel.Fail);
                }
            }
            catch (Exception ex) {
                Log($"Error deleting registry key: \"HKEY_LOCAL_MACHINE\\{key}\". Exception: {ex.Message}", LogLevel.Error);
            }
        }
    }

    /// <summary>
    /// Deletes each subkey from the given parent
    /// </summary>
    private static void DeleteSubKeys(RegistryKey parentKey) {
        foreach (var subKeyName in parentKey.GetSubKeyNames()) {
            using var subKey = parentKey.OpenSubKey(subKeyName, true);
            if (subKey is not null) {
                DeleteSubKeys(subKey);
                parentKey.DeleteSubKeyTree(subKeyName);
            }
        }
    }

    /// <summary>
    /// Finds all services that has a display name of "metallic" and runs "sc delete {service_name}" to delete the service.
    /// </summary>
    private static async Task DeleteServices(string[] expectedserviceNames) {
        var metalicServices = ServiceController.GetServices().Where(d => d.DisplayName.StartsWith("metallic", StringComparison.CurrentCultureIgnoreCase)).ToList();

        foreach (var service in metalicServices) {
            var serviceName = service.ServiceName;

            if (expectedserviceNames.Any(x => x.Equals(serviceName))) {
                foundServicesToCheckLater.Add(serviceName);
                Process.Start("sc", $"delete {serviceName}");
                Log($"Deleted service: {serviceName}", LogLevel.Success);
                await Task.Delay(TimeSpan.FromSeconds(2));
                continue;
            }
            else {
                Log($"Unable to find expected service name for {serviceName}", LogLevel.Fail);
                continue;
            }
        }
    }

    /// <summary>
    /// Check and compare the old list of services to the new current set
    /// </summary>
    private static void CheckNewlyInstalledServices() {
        var currentMetalicServices = ServiceController.GetServices().Where(d => d.DisplayName.Contains("metallic", StringComparison.CurrentCultureIgnoreCase)).ToList();
        var currentSetOfMetallicServices = new List<string>();

        foreach (var newService in currentMetalicServices)
            currentSetOfMetallicServices.Add(newService.ServiceName);

        foreach (var item in currentSetOfMetallicServices.Intersect(foundServicesToCheckLater))
            Log($"Service: \"{item}\" was successfully reinstalled", LogLevel.Success);

        foreach (var item in foundServicesToCheckLater.Except(currentSetOfMetallicServices))
            Log($"Service: \"{item}\" was not found in current set of services. This is expected.", LogLevel.Success);

        foreach (var item in currentSetOfMetallicServices.Except(foundServicesToCheckLater))
            Log($"Service: \"{item}\" is new.", LogLevel.Warning);
    }

    /// <summary>
    /// Copies a file from one place to another
    /// </summary>
    private static void CopyFile(string source, string destination) {
        var file = Path.GetFileName(source);
        Log($"Copying file: \"{file}\", please wait...");
        try {
            if (File.Exists(source)) {
                File.Copy(source, destination, true);
                Log($"Copied file: \"{file}\" to \"{destination}\"");
            }
        }
        catch (FileNotFoundException e) {
            Log($"Either \"{source}\" or \"{destination}\" is not available. Please check the paths manually and let Lily know about this.", LogLevel.Error);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(e.Message + "\n" + e.StackTrace);
            Console.ResetColor();
        }
    }

    /// <summary>
    /// Executes an application
    /// </summary>
    private static async Task ExecuteFile(string filePath) {
        if (File.Exists(filePath)) {
            Log($"Executing file: \"{filePath}\"");
            var process = Process.Start(new ProcessStartInfo() {
                FileName = filePath,
                UseShellExecute = true,
                Verb = "runas"
            });

            windows_Server64 = new FileInfo(filePath);

            Log($"Waiting for \"{windows_Server64.Name}\" to exit...");
            process?.WaitForExit();

            await Task.Delay(TimeSpan.FromSeconds(30)); // force a wait incase server is running slower
            CheckForSetupProcess();
        }
    }

    /// <summary>
    /// Hold current process until the Commvault Setup.exe is finished
    /// </summary>
    private static void CheckForSetupProcess() {
        Log("Waiting for the Commvault Setup to finish...");
        while (Process.GetProcesses().Any(p => p.ProcessName.Equals("Setup"))) {
            Thread.Sleep(TimeSpan.FromSeconds(1)); // wait and check again every one second
        }
    }

    /// <summary>
    /// Opens the specified file with Notepad
    /// </summary>
    private static void OpenNotepadForKey(string filePath) => Process.Start("notepad.exe", filePath);

    /// <summary>
    /// Remove the large file and close notepad
    /// </summary>
    private static void CleanupFiles() {
        if (windows_Server64 is not null) {
            Console.WriteLine();
            Log($"Deleting \"{windows_Server64.Name}\"...");
            windows_Server64.Delete();
            Log("Complete");
        }
        Console.WriteLine();
        Log("Closing Notepad");
        var notepadProcess = Process.GetProcesses().Where(x => x.ProcessName.Equals("notepad")).FirstOrDefault();
        if (notepadProcess is not null) {
            notepadProcess!.Kill();
            Console.WriteLine();
        }
    }

    /// <summary>
    /// Checks to see if current process is being ran as admin
    /// </summary>
    private static bool IsAdministrator() => new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);

    /// <summary>
    /// Console log output
    /// </summary>
    private static void Log(string message, LogLevel logLevel = LogLevel.Info) {
        Console.Write("[");
        Console.Write(DateTime.Now);
        Console.Write("] [");

        switch (logLevel) {
            case LogLevel.Warning:
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write("WARN");
                break;
            case LogLevel.Fail:
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.Write("FAIL");
                break;
            case LogLevel.Success:
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write("GOOD");
                break;
            case LogLevel.Error:
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("ERR-");
                break;
            case LogLevel.Info:
            default:
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write("INFO");
                break;
        }

        Console.ResetColor();
        Console.WriteLine("] " + message);
        Console.ResetColor();
    }
}

public enum LogLevel {
    Info,    // 0
    Warning, // 1
    Fail,    // 2
    Success, // 3
    Error    // 4
}