using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Setup.Configuration;
using System.Diagnostics;
using Vitevic.Shared.Extensions;

namespace Vitevic.VSInstaller
{
    [DebuggerDisplay("{DisplayName}")]
    sealed class VsInstance : IEquatable<VsInstance>, IComparable<VsInstance>
    {
        private const int Vs2015Major = 14;
        private const string Vs2015MajorStr = "14";
        private const int Vs2017Major = 15;
        private const int Vs2019Major = 16;

        /// <summary>
        /// Full installation version (like "15.9.28307.222").
        /// </summary>
        public Version Version { get; }

        /// <summary>
        /// Full display name (like "Visual Studio Professional 2017").
        /// </summary>
        public string DisplayName { get; }

        /// <summary>
        /// Visual Studio Edition (like "VsEdition.Professional")
        /// </summary>
        VsEdition Edition { get; }

        /// <summary>
        /// Path to the root product directory (like "C:\Program Files (x86)\Microsoft Visual Studio 14.0\").
        /// Always contains trailing slash.
        /// </summary>
        public string ProductDirectory { get; }

        /// <summary>
        /// The directory with devenv.exe, usually "$(ProductDir)Common7\IDE".
        /// </summary>
        public string EnvironmentDirectory { get; }

        /// <summary>
        /// Full path to "devenv.exe".
        /// </summary>
        public string EnvironmentPath { get; }

        /// <summary>
        /// A property indicating whether the instance is a prerelease (preview).
        /// </summary>
        public bool IsPrerelease { get; }

        /// <summary>
        /// Instance nickname given during the installation or empty string.
        /// For Vs2015 it is always empty string.
        /// </summary>
        public string Nickname { get; }

        /// <summary>
        /// The instance identifier (matches the name of the parent instance directory).
        /// For Vs2015 it is empty string.
        /// </summary>
        public string Id { get; }

        /// <summary>
        /// Root suffix (like "Exp").
        /// </summary>
        public string RootSuffix { get; }


        public bool Is2015
        {
            get
            {
                return Version.Major == Vs2015Major;
            }
        }

        public bool Is2017
        {
            get
            {
                return Version.Major == Vs2017Major;
            }
        }

        public bool Is2019
        {
            get
            {
                return Version.Major == Vs2019Major;
            }
        }

        #region IEquatable
        public bool Equals(VsInstance other)
        {
            // TODO: StringPath.Equals for directories
            return !(other == null)
                && string.Equals(RootSuffix, other.RootSuffix, StringComparison.OrdinalIgnoreCase)
                && string.Equals(ProductDirectory, other.ProductDirectory, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as VsInstance);
        }

        public override int GetHashCode()
        {
            return StringComparer.OrdinalIgnoreCase.GetHashCode(ProductDirectory)
                 ^ StringComparer.OrdinalIgnoreCase.GetHashCode(RootSuffix);
        }

        public static bool operator ==(VsInstance left, VsInstance right)
        {
            return object.Equals(left, right);
        }

        public static bool operator !=(VsInstance left, VsInstance right)
        {
            return !(left == right);
        }

        #endregion IEquatable

        #region IComparable
        public int CompareTo(VsInstance other)
        {
            if (other == null)
            {
                return 1;
            }

            var res = Version.CompareTo(other.Version);
            if( res == 0 )
            {
                // same version, for VS2017+ might be different instances
                res = StringComparer.OrdinalIgnoreCase.Compare(ProductDirectory, other.ProductDirectory);
            }

            if ( res == 0 )
            {
                res = StringComparer.OrdinalIgnoreCase.Compare(RootSuffix, other.RootSuffix);
            }

            return res;
        }

        public static bool operator <(VsInstance left, VsInstance right)
        {
            return right != null && right.CompareTo(left) > 0;
        }

        public static bool operator >(VsInstance left, VsInstance right)
        {
            return left != null && left.CompareTo(right) > 0;
        }

        public static bool operator <=(VsInstance left, VsInstance right)
        {
            return !(left > right);
        }

        public static bool operator >=(VsInstance left, VsInstance right)
        {
            return !(left < right);
        }
        #endregion

        VsInstance(Version version, string displayName, VsEdition edition, string productDir, string envDir, string envPath, bool prerelease, string nickname, string id, string rootSuffix)
        {
            Version = version;
            DisplayName = displayName;
            Edition = edition;
            ProductDirectory = productDir;
            EnvironmentDirectory = envDir;
            EnvironmentPath = envPath;
            IsPrerelease = prerelease;
            Nickname = nickname ?? "";
            Id = id ?? "";
            RootSuffix = rootSuffix ?? "";
        }

        public static IReadOnlyList<VsInstance> PopulateAll(int majorOldest = Vs2015Major)
        {
            HashSet<VsInstance> set = new HashSet<VsInstance>();
            // at first look for VS2015
            LookForPre2017(set);
            LookFor2017AndLater(set);

            return set.Where(x => x.Version.Major >= majorOldest).OrderBy(x => x).ToList();
        }

        private static void LookForPre2017(ISet<VsInstance> set)
        {
            using (var root = RegistryHive.LocalMachine.Registry32(@"Software\Microsoft\VisualStudio"))
            {
                var versions = root.GetSubKeyNames();
                foreach (var name in versions)
                {
                    var parts = name.Split('.');
                    if (parts.Length == 2 && Version.TryParse(name, out var version) )
                    {
                        // we are only interested in 14.*
                        using (var key = root.OpenSubKey(name + @"\Setup\vs"))
                        {
                            if( key != null )
                            {
                                var productDir = (string)key.GetValue("ProductDir");
                                var envDir = (string)key.GetValue("EnvironmentDirectory");
                                var envPath = (string)key.GetValue("EnvironmentPath");
                                var edition = GetPre2017Edition(key);
                                var editionStr = "";
                                if ( edition != VsEdition.None)
                                {
                                    editionStr = " " + edition;
                                }
                                var yearName = GetYearNameByMajorVersion(parts[0]);
                                var displayName = $"Visual Studio{editionStr} {yearName}";
                                var instance = new VsInstance(version, displayName, edition, productDir, envDir, envPath, false, null, null, null);
                                set.Add(instance);
                            }
                        }
                    }
                }
            }
        }

        private static VsEdition GetPre2017Edition(RegistryKey key)
        {
            using (var community = key.OpenSubKey("community"))
            {
                if (community != null)
                    return VsEdition.Community;
            }

            using (var pro = key.OpenSubKey("pro"))
            {
                if (pro != null)
                    return VsEdition.Professional;
            }

            return VsEdition.None;
        }

        private static string GetYearNameByMajorVersion(string v)
        {
            switch(v)
            {
                case "10": return "2010";
                case "11": return "2012";
                case "12": return "2013";
                case "14": return "2015";
            }

            return string.Empty;
        }

        private static void LookFor2017AndLater(HashSet<VsInstance> set)
        {
            const int REGDB_E_CLASSNOTREG = unchecked((int)0x80040154);

            try
            {
                var query = new SetupConfiguration();
                var query2 = (ISetupConfiguration2)query;
                // var e = query2.EnumAllInstances();
                var e = query2.EnumInstances();
                // var helper = (ISetupHelper)query;

                int fetched;
                var instances = new ISetupInstance[1];
                do
                {
                    e.Next(1, instances, out fetched);
                    if (fetched > 0)
                    {
                        var instance2 = (ISetupInstance2)instances[0];
                        var productDir = instance2.GetInstallationPath();
                        var productPath = instance2.GetProductPath();
                        var envPath = System.IO.Path.Combine(productDir, productPath);
                        var envDir = System.IO.Path.GetDirectoryName(envPath);
                        var version = Version.Parse(instance2.GetInstallationVersion());
                        var product = instance2.GetProduct();
                        var editionStr = product.GetId();
                        editionStr = editionStr.Substring(editionStr.LastIndexOf('.')+1);
                        var edition = (VsEdition)Enum.Parse(typeof(VsEdition), editionStr, true);
                        var displayName = instance2.GetDisplayName(1033);

                        var nickname = "";
                        var id = instance2.GetInstanceId();
                        var prerelease = false;
                        var catalog = instance2 as ISetupInstanceCatalog;
                        if( catalog != null)
                        {
                            prerelease = catalog.IsPrerelease();
                        }
                        var props = instance2.GetProperties();
                        if( props.GetNames().Contains("nickname") )
                        {
                            nickname = props.GetValue("nickname");
                        }

                        var instance = new VsInstance(version, displayName, edition, productDir, envDir, envPath, prerelease, nickname, id, null);
                        set.Add(instance);
                    }
                }
                while (fetched > 0);
            }
            catch (COMException ex) when (ex.HResult == REGDB_E_CLASSNOTREG)
            {
                Debug.WriteLine("The VS query API is not registered. Assuming no VS2017+ instances are installed.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error 0x{ex.HResult:x8}: {ex.Message}");
            }
        }
    }
}
