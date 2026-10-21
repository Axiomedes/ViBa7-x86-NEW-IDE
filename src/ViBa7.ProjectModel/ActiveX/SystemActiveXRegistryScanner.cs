using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Versioning;
using Microsoft.Win32;

namespace ViBa7.ProjectModel.ActiveX;

/// <summary>
/// Información de un componente ActiveX (OCX o DLL) registrado en el sistema operativo.
/// </summary>
public class RegisteredActiveXComponentInfo
{
    public string Guid { get; set; } = string.Empty;
    public int MajorVersion { get; set; } = 1;
    public int MinorVersion { get; set; } = 0;
    public string Lcid { get; set; } = "0";
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string FilePath { get; set; } = string.Empty;
    public bool IsControl { get; set; }
    public string FileExtension { get; set; } = string.Empty;
    public bool FileExists { get; set; } = true;

    public string DisplayTitle
    {
        get
        {
            string title = !string.IsNullOrWhiteSpace(Description) ? Description : Name;
            if (string.IsNullOrWhiteSpace(title))
            {
                title = Path.GetFileName(FilePath);
            }
            return title;
        }
    }
}

/// <summary>
/// Escáner del Registro de Windows para descubrir todos los controles ActiveX (.ocx)
/// y DLLs ActiveX / TypeLibs registradas en el sistema (32 y 64 bits), con capacidad
/// de autoregistro en el sistema.
/// </summary>
[SupportedOSPlatform("windows")]
public static class SystemActiveXRegistryScanner
{
    private const int REGKIND_REGISTER = 1;
    private const int REGKIND_NONE = 2;

    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void LoadTypeLibEx(
        string szFile,
        int regkind,
        out ITypeLib pptlib);

    /// <summary>
    /// Enumera todos los componentes ActiveX (Controles OCX y DLLs ActiveX / TypeLibs)
    /// registrados en el Registro de Windows en las vistas de 32 bits y 64 bits.
    /// </summary>
    public static List<RegisteredActiveXComponentInfo> GetRegisteredComponents()
    {
        var result = new Dictionary<string, RegisteredActiveXComponentInfo>(StringComparer.OrdinalIgnoreCase);

        // 1. Escanear TypeLib en HKCR (32-bit y 64-bit)
        ScanTypeLibRegistryHive(RegistryHive.ClassesRoot, RegistryView.Registry32, result);
        ScanTypeLibRegistryHive(RegistryHive.ClassesRoot, RegistryView.Registry64, result);

        // 2. Escanear HKLM\SOFTWARE\Classes\TypeLib y HKLM\SOFTWARE\WOW6432Node\Classes\TypeLib
        ScanTypeLibRegistryHive(RegistryHive.LocalMachine, RegistryView.Registry32, result, @"SOFTWARE\Classes\TypeLib");
        ScanTypeLibRegistryHive(RegistryHive.LocalMachine, RegistryView.Registry64, result, @"SOFTWARE\Classes\TypeLib");
        ScanTypeLibRegistryHive(RegistryHive.LocalMachine, RegistryView.Registry64, result, @"SOFTWARE\WOW6432Node\Classes\TypeLib");

        // 3. Complementar con Controles clásicos presentes en SysWOW64 y System32
        ScanStandardWindowsDirectories(result);

        return result.Values
            .OrderBy(c => c.IsControl ? 0 : 1) // Controles primero
            .ThenBy(c => c.DisplayTitle, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    private static void ScanTypeLibRegistryHive(
        RegistryHive hive,
        RegistryView view,
        Dictionary<string, RegisteredActiveXComponentInfo> result,
        string subKeyPath = "TypeLib")
    {
        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(hive, view);
            using var typeLibKey = baseKey.OpenSubKey(subKeyPath);
            if (typeLibKey == null) return;

            foreach (var guidName in typeLibKey.GetSubKeyNames())
            {
                if (string.IsNullOrWhiteSpace(guidName) || !guidName.StartsWith('{')) continue;

                using var guidKey = typeLibKey.OpenSubKey(guidName);
                if (guidKey == null) continue;

                foreach (var verName in guidKey.GetSubKeyNames())
                {
                    using var verKey = guidKey.OpenSubKey(verName);
                    if (verKey == null) continue;

                    string description = verKey.GetValue(null) as string ?? string.Empty;

                    int major = 1;
                    int minor = 0;
                    string[] verParts = verName.Split('.');
                    if (verParts.Length > 0 && int.TryParse(verParts[0], out int maj)) major = maj;
                    if (verParts.Length > 1 && int.TryParse(verParts[1], out int min)) minor = min;

                    // Buscar ruta del archivo en subclaves win32 o win64
                    string? filePath = FindFilePathInVersionKey(verKey);
                    if (string.IsNullOrWhiteSpace(filePath)) continue;

                    filePath = CleanPath(filePath);
                    if (string.IsNullOrWhiteSpace(filePath)) continue;

                    string ext = Path.GetExtension(filePath).ToLowerInvariant();
                    // Solo considerar extensiones de componentes binarios ActiveX y librerías de tipos
                    if (ext is not (".ocx" or ".dll" or ".tlb" or ".olb" or ".exe")) continue;

                    // Determinar si es control
                    bool isControl = ext == ".ocx";
                    if (!isControl)
                    {
                        // Verificar si tiene la marca FLAGS o subclave CONTROL
                        using var flagsKey = verKey.OpenSubKey("FLAGS");
                        if (flagsKey != null)
                        {
                            object? flagsVal = flagsKey.GetValue(null);
                            if (flagsVal is int f && (f & 2) != 0) isControl = true; // LIBFLAG_FCONTROL = 2
                        }
                    }

                    string key = $"{guidName}_{major}.{minor}_{filePath}";
                    if (!result.ContainsKey(key))
                    {
                        result[key] = new RegisteredActiveXComponentInfo
                        {
                            Guid = guidName.ToUpperInvariant(),
                            MajorVersion = major,
                            MinorVersion = minor,
                            Description = description,
                            Name = !string.IsNullOrWhiteSpace(description) ? description : Path.GetFileNameWithoutExtension(filePath),
                            FilePath = filePath,
                            IsControl = isControl,
                            FileExtension = ext,
                            FileExists = File.Exists(filePath)
                        };
                    }
                }
            }
        }
        catch
        {
            // Tolerar excepciones de permisos de registro
        }
    }

    private static string? FindFilePathInVersionKey(RegistryKey verKey)
    {
        // 1. Estructura estándar: {ver}\0\win32 o {ver}\0\win64
        foreach (var sub in verKey.GetSubKeyNames())
        {
            using var lcidKey = verKey.OpenSubKey(sub);
            if (lcidKey == null) continue;

            using var win32Key = lcidKey.OpenSubKey("win32");
            if (win32Key?.GetValue(null) is string p32 && !string.IsNullOrWhiteSpace(p32))
            {
                return p32;
            }

            using var win64Key = lcidKey.OpenSubKey("win64");
            if (win64Key?.GetValue(null) is string p64 && !string.IsNullOrWhiteSpace(p64))
            {
                return p64;
            }
        }

        // 2. Estructura directa {ver}\win32
        using var direct32 = verKey.OpenSubKey("win32");
        if (direct32?.GetValue(null) is string directPath && !string.IsNullOrWhiteSpace(directPath))
        {
            return directPath;
        }

        using var direct64 = verKey.OpenSubKey("win64");
        if (direct64?.GetValue(null) is string directPath64 && !string.IsNullOrWhiteSpace(directPath64))
        {
            return directPath64;
        }

        return null;
    }

    private static void ScanStandardWindowsDirectories(Dictionary<string, RegisteredActiveXComponentInfo> result)
    {
        var dirs = new List<string>();

        string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        string sysWow64 = Path.Combine(winDir, "SysWOW64");
        string sys32 = Environment.SystemDirectory;

        if (Directory.Exists(sysWow64)) dirs.Add(sysWow64);
        if (Directory.Exists(sys32) && !string.Equals(sys32, sysWow64, StringComparison.OrdinalIgnoreCase)) dirs.Add(sys32);

        foreach (var dir in dirs)
        {
            try
            {
                var ocxFiles = Directory.GetFiles(dir, "*.ocx");
                foreach (var ocx in ocxFiles)
                {
                    string fileName = Path.GetFileName(ocx);
                    if (result.Values.Any(r => string.Equals(Path.GetFileName(r.FilePath), fileName, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    // Intentar leer la metadata básica de la TypeLib
                    if (OcxTypeLibReader.TryReadTypeLib(ocx, out var lib, out _))
                    {
                        string key = $"{lib!.Guid}_{lib.MajorVersion}.{lib.MinorVersion}_{ocx}";
                        if (!result.ContainsKey(key))
                        {
                            result[key] = new RegisteredActiveXComponentInfo
                            {
                                Guid = lib.Guid,
                                MajorVersion = lib.MajorVersion,
                                MinorVersion = lib.MinorVersion,
                                Description = !string.IsNullOrWhiteSpace(lib.Description) ? lib.Description : lib.LibraryName,
                                Name = lib.LibraryName,
                                FilePath = ocx,
                                IsControl = true,
                                FileExtension = ".ocx",
                                FileExists = true
                            };
                        }
                    }
                }
            }
            catch
            {
                // Ignorar errores de acceso
            }
        }
    }

    private static string CleanPath(string rawPath)
    {
        string p = rawPath.Trim();
        if (p.StartsWith('\"') && p.EndsWith('\"') && p.Length >= 2)
        {
            p = p[1..^1].Trim();
        }

        // Expandir variables de entorno como %SystemRoot%
        p = Environment.ExpandEnvironmentVariables(p);

        // Si la ruta incluye sufijos de índice de recurso COM (e.g. "C:\Path\foo.dll\1" o "foo.dll\2")
        int lastSlash = p.LastIndexOfAny(['\\', '/']);
        if (lastSlash >= 0)
        {
            string fileNamePart = p[(lastSlash + 1)..];
            int subSlash = fileNamePart.IndexOfAny(['\\', '/']);
            if (subSlash > 0)
            {
                p = p[..(lastSlash + 1 + subSlash)];
            }
            else
            {
                // Verificar si termina en \digito
                int digitSlash = p.LastIndexOf('\\');
                if (digitSlash > 0 && digitSlash == p.Length - 2 && char.IsDigit(p[^1]))
                {
                    p = p[..digitSlash];
                }
            }
        }

        return p;
    }

    /// <summary>
    /// Determina si un binario PE es de 32 bits (x86) inspeccionando su cabecera PE.
    /// </summary>
    private static bool Is32BitBinary(string filePath)
    {
        try
        {
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var br = new BinaryReader(fs);
            if (fs.Length < 64) return true; // Asumir 32 bits por defecto para componentes clásicos de VB6
            if (br.ReadUInt16() != 0x5A4D) return true; // 'MZ'
            fs.Seek(0x3C, SeekOrigin.Begin);
            int peOffset = br.ReadInt32();
            if (peOffset <= 0 || peOffset + 6 > fs.Length) return true;
            fs.Seek(peOffset, SeekOrigin.Begin);
            if (br.ReadUInt32() != 0x00004550) return true; // 'PE\0\0'
            ushort machine = br.ReadUInt16();
            return machine == 0x014C; // IMAGE_FILE_MACHINE_I386
        }
        catch
        {
            return true;
        }
    }

    /// <summary>
    /// Autoregistra un archivo .ocx o .dll en el sistema operativo mediante
    /// la instancia adecuada de regsvr32.exe (SysWOW64 para 32-bit o System32 para 64-bit)
    /// y gestiona permisos de elevación si son requeridos.
    /// </summary>
    public static bool RegisterActiveXFileInSystem(string filePath, out string? statusMessage)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        if (!File.Exists(filePath))
        {
            statusMessage = $"El archivo no existe en disco: {filePath}";
            return false;
        }

        // 1. Determinar el ejecutable regsvr32 correspondiente a la arquitectura del binario
        string regsvrPath;
        if (Environment.Is64BitOperatingSystem && Is32BitBinary(filePath))
        {
            regsvrPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SysWOW64", "regsvr32.exe");
            if (!File.Exists(regsvrPath)) regsvrPath = "regsvr32.exe";
        }
        else
        {
            regsvrPath = Path.Combine(Environment.SystemDirectory, "regsvr32.exe");
            if (!File.Exists(regsvrPath)) regsvrPath = "regsvr32.exe";
        }

        // 2. Intentar ejecutar regsvr32 silencioso sin elevación
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = regsvrPath,
                Arguments = $"/s \"{filePath}\"",
                UseShellExecute = true,
                CreateNoWindow = true
            };

            using var proc = Process.Start(psi);
            if (proc != null)
            {
                proc.WaitForExit(6000);
                if (proc.ExitCode == 0)
                {
                    statusMessage = "Componente autoregistrado con éxito en el sistema.";
                    return true;
                }
            }
        }
        catch
        {
            // Reintentar con elevación
        }

        // 3. Si falló el registro estándar (posible falta de permisos de Administrador), reintentar con elevación (runas)
        try
        {
            var psiElevated = new ProcessStartInfo
            {
                FileName = regsvrPath,
                Arguments = $"/s \"{filePath}\"",
                UseShellExecute = true,
                Verb = "runas"
            };

            using var procElevated = Process.Start(psiElevated);
            if (procElevated != null)
            {
                procElevated.WaitForExit(8000);
                if (procElevated.ExitCode == 0)
                {
                    statusMessage = "Componente autoregistrado con éxito con permisos de administrador.";
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            statusMessage = $"El registro requiere permisos de administrador o falló: {ex.Message}";
            return false;
        }

        statusMessage = "El comando de registro finalizó con código de error.";
        return false;
    }
}
