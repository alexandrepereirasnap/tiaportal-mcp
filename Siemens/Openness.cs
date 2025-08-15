
using Microsoft.Win32;
using Siemens.Collaboration.Net;
using Siemens.Engineering;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Threading.Tasks;

namespace TiaMcpServer.Siemens
{
    public static class Openness
    {
        public static int TiaMajorVersion { get; private set; }

        public static void Initialize(int? tiaMajorVersion = 20)
        {
            // This initialization ensures compatibility with both future and past TIA Portal Openness NuGet packages.
            // Specifically, it addresses scenarios related to:
            // 2.1 Siemens.Collaboration.Net.TiaPortal.Openness.Resolver (requires TiaPortalLocation environment variable)
            // 2.2 Siemens.Collaboration.Net.TiaPortal.Packages.Openness
            // 2.3 Direct API initialization: Api.Global.Openness().Initialize(tiaMajorVersion: 20);

            TiaMajorVersion = tiaMajorVersion ?? 20; // Default to TIA Portal V20 if not specified


            AppDomain.CurrentDomain.AssemblyResolve += MyResolver;
        }

        public static async Task<bool> IsUserInGroup()
        {
            if (Api.Global.Openness().IsUserInGroup())
            {
                // user is in group
                return true;
            }
            else
            {
                return await Api.Global.Openness().AddUserToGroupAsync();
            }
        }


        private static Assembly MyResolver(object sender, ResolveEventArgs args)
        {
            var assemblyName = new AssemblyName(args.Name);
            if (!assemblyName.Name.StartsWith("Siemens.Engineering"))
            {
                return null;
            }

            var tiaInstallPath = GetTiaPortalInstallPath();
            if (string.IsNullOrEmpty(tiaInstallPath))
            {
                throw new InvalidOperationException($"Could not find TIA Portal installation path for version {TiaMajorVersion} in the registry.");
            }

            var majorVersionString = TiaMajorVersion.ToString();
            var searchDirectories = new[] 
            {
                Path.Combine(tiaInstallPath, "PublicAPI", $"V{majorVersionString}"),
                Path.Combine(tiaInstallPath, "Bin", "PublicAPI")
            };

            var versionsToIgnore = new[] { "V13", "V14", "V15", "V16", "V17", "V18", "V19", "V20" }
                                    .Where(v => v != $"V{majorVersionString}");

            foreach (var dir in searchDirectories)
            {
                var assemblyPath = FindAssemblyRecursive(dir, assemblyName.Name + ".dll", versionsToIgnore);
                if (assemblyPath != null)
                {
                    return Assembly.LoadFrom(assemblyPath);
                }
            }

            throw new FileNotFoundException($"Could not find DLL '{assemblyName.Name}' for TIA Portal version {TiaMajorVersion} in the installation directories.");
        }

        private static string GetTiaPortalInstallPath()
        {
            var subKeyName = $@"SOFTWARE\Siemens\Automation\_InstalledSW\TIAP{TiaMajorVersion}\TIA_Opns";

            using (var regBaseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            using (var tiaOpnsKey = regBaseKey.OpenSubKey(subKeyName))
            {
                return tiaOpnsKey?.GetValue("Path")?.ToString();
            }
        }

        private static string FindAssemblyRecursive(string directory, string fileName, IEnumerable<string> excludedDirectories)
        {
            if (!Directory.Exists(directory))
            {
                return null;
            }

            var filePath = Path.Combine(directory, fileName);
            if (File.Exists(filePath))
            {
                return filePath;
            }

            foreach (var subDir in Directory.GetDirectories(directory))
            {
                var subDirName = new DirectoryInfo(subDir).Name;
                if (excludedDirectories.Contains(subDirName))
                {
                    continue;
                }

                var result = FindAssemblyRecursive(subDir, fileName, excludedDirectories);
                if (result != null)
                {
                    return result;
                }
            }

            return null;
        }
    }
}
