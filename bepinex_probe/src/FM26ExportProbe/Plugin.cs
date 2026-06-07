using System;
using System.IO;
using BepInEx;
using BepInEx.Logging;
using BepInEx.Unity.IL2CPP;
using Il2CppInterop.Runtime.Injection;
using UnityEngine;

namespace FM26ExportProbe;

[BepInPlugin(PluginInfo.GUID, PluginInfo.NAME, PluginInfo.VERSION)]
public class Plugin : BasePlugin
{
    internal static ManualLogSource Log { get; private set; } = null!;
    internal static string ExportDirectory { get; private set; } = "";

    public override void Load()
    {
        Log = base.Log;
        ExportDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "FM26Exports"
        );

        Log.LogInfo($"{PluginInfo.NAME} v{PluginInfo.VERSION} — loading");

        ClassInjector.RegisterTypeInIl2Cpp<ProbeBehaviour>();
        AddComponent<ProbeBehaviour>();

        Log.LogInfo($"{PluginInfo.NAME} behaviour registered");
    }
}

public static class PluginInfo
{
    public const string GUID = "com.local.fm26.exportprobe";
    public const string NAME = "FM26 Export Probe";
    public const string VERSION = "0.1.0";
}

public class ProbeBehaviour : MonoBehaviour
{
    private bool _initialized;

    public ProbeBehaviour(IntPtr ptr) : base(ptr) { }

    private void Awake()
    {
        try
        {
            Directory.CreateDirectory(Plugin.ExportDirectory);

            var stamp = Path.Combine(Plugin.ExportDirectory, "probe_loaded.txt");
            File.WriteAllText(
                stamp,
                $"FM26 Export Probe loaded at {DateTime.UtcNow:O}{Environment.NewLine}"
            );

            Plugin.Log.LogInfo("FM26 Export Probe loaded");
            Plugin.Log.LogInfo($"Export directory: {Plugin.ExportDirectory}");
            Plugin.Log.LogInfo($"Wrote {stamp}");

            _initialized = true;
        }
        catch (Exception ex)
        {
            Plugin.Log.LogError($"Probe Awake failed: {ex}");
        }
    }

    private void Update()
    {
        if (!_initialized)
        {
            return;
        }

        if (!Input.GetKeyDown(KeyCode.F8))
        {
            return;
        }

        try
        {
            Plugin.Log.LogInfo("F8 pressed — writing diagnostics");
            var path = Exporter.WriteDiagnostics(Plugin.ExportDirectory);
            Plugin.Log.LogInfo($"Diagnostics written to {path}");
        }
        catch (Exception ex)
        {
            Plugin.Log.LogError($"F8 diagnostic export failed: {ex}");
        }
    }
}
