using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using UnityEngine;
using UnityEngine.SceneManagement;

namespace FM26ExportProbe;

/// <summary>
/// Diagnostic exporter only for Stage 4 research.
/// TODO: identify visible player list rows/cells and serialize them.
/// </summary>
public static class Exporter
{
    public static string WriteDiagnostics(string exportDirectory)
    {
        Directory.CreateDirectory(exportDirectory);
        var path = Path.Combine(exportDirectory, "fm26_probe_diagnostic.txt");

        var sb = new StringBuilder();
        sb.AppendLine("FM26 Export Probe — diagnostic snapshot");
        sb.AppendLine($"UTC: {DateTime.UtcNow:O}");
        sb.AppendLine($"Unity version: {Application.unityVersion}");
        sb.AppendLine($"Platform: {Application.platform}");
        sb.AppendLine($"Product: {Application.productName}");
        sb.AppendLine();

        AppendAssemblies(sb);
        AppendScenes(sb);
        sb.AppendLine();
        sb.Append(UiScanner.Scan());

        sb.AppendLine();
        sb.AppendLine("=== Export status ===");
        sb.AppendLine("Player table extraction: NOT IMPLEMENTED (diagnostics only).");
        sb.AppendLine("Next step: confirm probe load, then reverse-engineer UI table structure.");

        File.WriteAllText(path, sb.ToString());
        return path;
    }

    private static void AppendAssemblies(StringBuilder sb)
    {
        sb.AppendLine("=== Loaded assemblies (sample) ===");
        var assemblies = AppDomain.CurrentDomain.GetAssemblies()
            .OrderBy(a => a.GetName().Name, StringComparer.OrdinalIgnoreCase)
            .Take(200)
            .ToList();

        foreach (var assembly in assemblies)
        {
            var name = assembly.GetName();
            sb.AppendLine($"  {name.Name} v{name.Version}");
        }

        if (AppDomain.CurrentDomain.GetAssemblies().Length > assemblies.Count)
        {
            sb.AppendLine($"  ... truncated ({AppDomain.CurrentDomain.GetAssemblies().Length} total)");
        }
    }

    private static void AppendScenes(StringBuilder sb)
    {
        sb.AppendLine();
        sb.AppendLine("=== Unity scenes ===");
        sb.AppendLine($"Active scene: {SceneManager.GetActiveScene().name}");
        sb.AppendLine($"Scene count: {SceneManager.sceneCount}");

        for (var i = 0; i < SceneManager.sceneCount; i++)
        {
            var scene = SceneManager.GetSceneAt(i);
            var roots = scene.isLoaded ? scene.GetRootGameObjects() : Array.Empty<GameObject>();
            sb.AppendLine($"  [{i}] {scene.name} — roots: {roots.Length}");
            foreach (var root in roots.Take(30))
            {
                sb.AppendLine($"      root: {root.name}");
            }
            if (roots.Length > 30)
            {
                sb.AppendLine($"      ... and {roots.Length - 30} more roots");
            }
        }
    }
}
