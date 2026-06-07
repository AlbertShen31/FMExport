using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UnityEngine;
using UnityEngine.SceneManagement;

namespace FM26ExportProbe;

/// <summary>
/// First-pass UI/scene scanner. Does not assume FM26-specific class names.
/// </summary>
public static class UiScanner
{
    private static readonly string[] LikelyTableKeywords =
    {
        "table", "grid", "row", "cell", "list", "squad", "player", "search", "browser",
    };

    public static string Scan()
    {
        var sb = new StringBuilder();
        sb.AppendLine("=== UI / Scene Scan ===");
        sb.AppendLine($"UTC: {DateTime.UtcNow:O}");
        sb.AppendLine();

        var sceneCount = SceneManager.sceneCount;
        sb.AppendLine($"Loaded scenes: {sceneCount}");

        var componentCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var keywordHits = new List<string>();

        for (var i = 0; i < sceneCount; i++)
        {
            var scene = SceneManager.GetSceneAt(i);
            sb.AppendLine();
            sb.AppendLine($"Scene [{i}]: name={scene.name}, path={scene.path}, loaded={scene.isLoaded}");

            if (!scene.isLoaded)
            {
                continue;
            }

            var roots = scene.GetRootGameObjects();
            sb.AppendLine($"  Root GameObjects: {roots.Length}");

            foreach (var root in roots)
            {
                AppendGameObjectTree(sb, root, depth: 1, componentCounts, keywordHits, maxDepth: 6);
            }
        }

        sb.AppendLine();
        sb.AppendLine("=== Component type counts (sampled) ===");
        foreach (var pair in componentCounts.OrderByDescending(p => p.Value).ThenBy(p => p.Key))
        {
            sb.AppendLine($"  {pair.Key}: {pair.Value}");
        }

        sb.AppendLine();
        sb.AppendLine("=== Likely table/grid/list nodes (name match) ===");
        if (keywordHits.Count == 0)
        {
            sb.AppendLine("  (none found by keyword)");
        }
        else
        {
            foreach (var hit in keywordHits.Distinct().Take(200))
            {
                sb.AppendLine($"  {hit}");
            }
            if (keywordHits.Count > 200)
            {
                sb.AppendLine($"  ... and {keywordHits.Count - 200} more");
            }
        }

        return sb.ToString();
    }

    private static void AppendGameObjectTree(
        StringBuilder sb,
        GameObject go,
        int depth,
        Dictionary<string, int> componentCounts,
        List<string> keywordHits,
        int maxDepth
    )
    {
        var indent = new string(' ', depth * 2);
        var active = go.activeInHierarchy ? "active" : "inactive";
        sb.AppendLine($"{indent}- {go.name} ({active})");

        var lowerName = go.name.ToLowerInvariant();
        if (LikelyTableKeywords.Any(k => lowerName.Contains(k)))
        {
            keywordHits.Add($"{go.name} @ depth {depth}");
        }

        var components = go.GetComponents<Component>();
        foreach (var component in components)
        {
            if (component == null)
            {
                continue;
            }

            var typeName = component.GetType().Name;
            componentCounts.TryGetValue(typeName, out var count);
            componentCounts[typeName] = count + 1;

            if (depth <= 3)
            {
                sb.AppendLine($"{indent}  [{typeName}]");
            }
        }

        if (depth >= maxDepth)
        {
            return;
        }

        var transform = go.transform;
        for (var i = 0; i < transform.childCount; i++)
        {
            AppendGameObjectTree(
                sb,
                transform.GetChild(i).gameObject,
                depth + 1,
                componentCounts,
                keywordHits,
                maxDepth
            );
        }
    }
}
