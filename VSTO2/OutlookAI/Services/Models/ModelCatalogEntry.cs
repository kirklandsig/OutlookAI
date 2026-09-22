using System;
using System.Collections.Generic;
using System.Linq;

namespace OutlookAI.Services.Models
{
    /// <summary>
    /// A model's upgrade hint from the catalog. With <see cref="RetirementAt"/>
    /// set it is a hard retirement (the model stops working then); without it,
    /// just a "newer model available" nudge.
    /// </summary>
    public sealed class ModelUpgradeInfo
    {
        public ModelUpgradeInfo(string model, DateTimeOffset? retirementAt)
        {
            Model = model;
            RetirementAt = retirementAt;
        }

        /// <summary>Suggested replacement slug; may be null.</summary>
        public string Model { get; }

        public DateTimeOffset? RetirementAt { get; }

        public bool IsRetiredAt(DateTimeOffset now)
        {
            return RetirementAt.HasValue && now >= RetirementAt.Value;
        }
    }

    /// <summary>One model from the Codex model catalog.</summary>
    public sealed class ModelCatalogEntry
    {
        public ModelCatalogEntry(
            string slug,
            string displayName,
            string description,
            bool listed,
            int priority,
            IEnumerable<string> efforts,
            ModelUpgradeInfo upgrade)
        {
            if (string.IsNullOrWhiteSpace(slug)) throw new ArgumentException("Model slug is required.", nameof(slug));
            Slug = slug;
            DisplayName = string.IsNullOrWhiteSpace(displayName) ? slug : displayName;
            Description = description ?? "";
            Listed = listed;
            Priority = priority;
            Efforts = (efforts ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            Upgrade = upgrade;
        }

        public string Slug { get; }
        public string DisplayName { get; }
        public string Description { get; }

        /// <summary>False for catalog models with visibility other than "list" (not shown in the picker).</summary>
        public bool Listed { get; }

        /// <summary>Catalog ordering; lower comes first.</summary>
        public int Priority { get; }

        /// <summary>Supported reasoning efforts as the catalog lists them (lowercase wire values).</summary>
        public IReadOnlyList<string> Efforts { get; }

        public ModelUpgradeInfo Upgrade { get; }
    }
}
