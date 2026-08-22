using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Windows.Input;
using MediaPipelineTray.Services;

namespace MediaPipelineTray.ViewModels;

/// <summary>
/// One editable setting.
///
/// A preset row can be inherited or overridden. Editing it creates an override; Reset removes
/// it so the value follows the global default again. That distinction is the whole point of
/// preset inheritance, so the UI shows it rather than flattening everything into a value.
/// </summary>
public sealed class SettingRow : Observable
{
    private string _value = "";
    private bool _isOverridden;

    public required SettingDefinition Definition { get; init; }

    /// <summary>Null for the global settings view.</summary>
    public string? Preset { get; init; }

    public required string InheritedValue { get; init; }

    public string Label => Definition.Label;
    public string Help => Definition.Help;
    public SettingKind Kind => Definition.Kind;
    public IReadOnlyList<string> Choices => Definition.Choices;

    public bool IsText => Kind is SettingKind.Integer or SettingKind.Decimal or SettingKind.Text;
    public bool IsChoice => Kind == SettingKind.Choice;
    public bool IsBoolean => Kind == SettingKind.Boolean;

    /// <summary>
    /// A preset option can only inherit if there is a global to inherit from. Options that
    /// exist solely per preset, like Grouping, fall back to the watcher's own default instead,
    /// so calling them "inherited" would be a lie.
    /// </summary>
    public bool CanInherit => Preset is not null && Definition.GlobalScoped;

    public string Value
    {
        get => _value;
        set
        {
            if (Set(ref _value, value))
            {
                IsOverridden = true;
                Raise(nameof(BoolValue));
                Changed?.Invoke(this, EventArgs.Empty);
            }
        }
    }

    public bool BoolValue
    {
        get => _value.Equals("true", StringComparison.OrdinalIgnoreCase)
               || _value == "1"
               || _value.Equals("yes", StringComparison.OrdinalIgnoreCase);
        set => Value = value ? "true" : "false";
    }

    public bool IsOverridden
    {
        get => _isOverridden;
        set
        {
            if (Set(ref _isOverridden, value))
            {
                Raise(nameof(InheritanceNote));
                Raise(nameof(ShowReset));
            }
        }
    }

    public bool ShowReset => CanInherit && IsOverridden;

    public string InheritanceNote => !CanInherit
        ? ""
        : IsOverridden
            ? $"overrides the default of {(InheritedValue.Length > 0 ? InheritedValue : "unset")}"
            : "inherited";

    public event EventHandler? Changed;

    public ICommand ResetCommand => _resetCommand ??= new RelayCommand(ResetToInherited);

    private ICommand? _resetCommand;

    public void ResetToInherited()
    {
        _value = InheritedValue;
        IsOverridden = false;
        Raise(nameof(Value));
        Raise(nameof(BoolValue));
        Changed?.Invoke(this, EventArgs.Empty);
    }
}

public sealed class SettingGroup
{
    public required string Name { get; init; }
    public required IReadOnlyList<SettingRow> Rows { get; init; }
}

/// <summary>Base for the two config-editing views, which differ only in scope.</summary>
public abstract class ConfigEditorViewModel : Observable
{
    private bool _isDirty;

    protected ConfigEditorViewModel(PipelinePaths paths, WatcherService watcher)
    {
        Paths = paths;
        Watcher = watcher;
    }

    protected PipelinePaths Paths { get; }
    protected WatcherService Watcher { get; }

    public ObservableCollection<SettingGroup> Groups { get; } = [];

    public bool IsDirty
    {
        get => _isDirty;
        protected set
        {
            if (Set(ref _isDirty, value))
            {
                Raise(nameof(DirtyNote));
            }
        }
    }

    public string DirtyNote => IsDirty
        ? "Unsaved changes. Saving restarts the watcher so they take effect."
        : "";

    public abstract void Load();

    public abstract void Save();

    protected void Track(SettingRow row) => row.Changed += (_, _) => IsDirty = true;

    /// <summary>
    /// Writes the file and restarts the watcher. Settings are only read at startup, so saving
    /// without restarting would leave the UI claiming something that is not true yet.
    /// </summary>
    public async Task<bool> SaveAndRestartAsync()
    {
        Save();
        IsDirty = false;

        if (!Watcher.IsRunning)
        {
            return true;
        }

        return await Watcher.RestartAsync(TimeSpan.FromSeconds(90)).ConfigureAwait(false);
    }
}

/// <summary>The global defaults every preset inherits.</summary>
public sealed class SettingsViewModel : ConfigEditorViewModel
{
    public SettingsViewModel(PipelinePaths paths, WatcherService watcher) : base(paths, watcher) { }

    public string ConfigPath => Paths.ConfigFile;
    public string PipelineRoot => Paths.PipelineRoot;

    public override void Load()
    {
        Groups.Clear();

        var ini = IniFile.Load(Paths.ConfigFile);
        var globals = ini.ReadGlobals();
        var definitions = SettingCatalog.ForGlobal().ToList();

        foreach (var group in SettingCatalog.Groups(definitions))
        {
            var rows = definitions
                .Where(d => d.Group == group)
                .Select(definition =>
                {
                    var row = new SettingRow
                    {
                        Definition = definition,
                        Preset = null,
                        InheritedValue = "",
                    };

                    row.Value = globals.TryGetValue(definition.Key, out var value) ? value : "";
                    row.IsOverridden = false;
                    Track(row);
                    return row;
                })
                .ToList();

            Groups.Add(new SettingGroup { Name = group, Rows = rows });
        }

        IsDirty = false;
    }

    public override void Save()
    {
        var ini = IniFile.Load(Paths.ConfigFile);

        foreach (var row in Groups.SelectMany(group => group.Rows))
        {
            if (row.Value.Length > 0)
            {
                ini.Set(row.Definition.Key, row.Value);
            }
        }

        ini.Save(Paths.ConfigFile);
    }
}

public sealed class PresetEditor : Observable
{
    public required string Name { get; init; }
    public required IReadOnlyList<SettingGroup> Groups { get; init; }
    public required string FolderPath { get; init; }

    public string Summary { get; set; } = "";
}

/// <summary>Per-preset options, plus adding and removing presets.</summary>
public sealed class PresetsViewModel : ConfigEditorViewModel
{
    private PresetEditor? _selected;

    public PresetsViewModel(PipelinePaths paths, WatcherService watcher) : base(paths, watcher) { }

    public ObservableCollection<PresetEditor> Presets { get; } = [];

    public PresetEditor? Selected
    {
        get => _selected;
        set
        {
            if (Set(ref _selected, value))
            {
                Raise(nameof(HasSelection));
            }
        }
    }

    public bool HasSelection => Selected is not null;

    public override void Load()
    {
        var previous = Selected?.Name;

        Presets.Clear();
        Groups.Clear();

        var ini = IniFile.Load(Paths.ConfigFile);
        var globals = ini.ReadGlobals();
        var presets = ini.ReadPresets();
        var definitions = SettingCatalog.ForPreset().ToList();

        foreach (var (name, overrides) in presets)
        {
            var groups = new List<SettingGroup>();

            foreach (var group in SettingCatalog.Groups(definitions))
            {
                var rows = definitions
                    .Where(d => d.Group == group)
                    .Select(definition =>
                    {
                        // A preset-only option has no global; its fallback is the watcher's default.
                        var inherited = definition.GlobalScoped && globals.TryGetValue(definition.Key, out var global)
                            ? global
                            : definition.Default;

                        var overridden = overrides.TryGetValue(definition.Key, out var value);

                        var row = new SettingRow
                        {
                            Definition = definition,
                            Preset = name,
                            InheritedValue = inherited,
                        };

                        row.Value = overridden ? value! : inherited;
                        row.IsOverridden = overridden;
                        Track(row);
                        return row;
                    })
                    .ToList();

                groups.Add(new SettingGroup { Name = group, Rows = rows });
            }

            Presets.Add(new PresetEditor
            {
                Name = name,
                Groups = groups,
                FolderPath = Path.Combine(Paths.PipelineRoot, name),
            });
        }

        Selected = Presets.FirstOrDefault(p => p.Name == previous) ?? Presets.FirstOrDefault();
        IsDirty = false;
    }

    public override void Save()
    {
        var ini = IniFile.Load(Paths.ConfigFile);

        foreach (var preset in Presets)
        {
            foreach (var row in preset.Groups.SelectMany(group => group.Rows))
            {
                if (row.IsOverridden && row.Value.Length > 0)
                {
                    ini.Set(row.Definition.Key, row.Value, preset.Name);
                }
                else
                {
                    ini.RemoveKey(row.Definition.Key, preset.Name);
                }
            }
        }

        ini.Save(Paths.ConfigFile);
    }

    /// <summary>
    /// Adds a preset with a minimal, working configuration rather than an empty section, so it
    /// does something sensible the moment the watcher restarts.
    /// </summary>
    public void AddPreset(string name)
    {
        var ini = IniFile.Load(Paths.ConfigFile);

        if (ini.HasPreset(name))
        {
            throw new InvalidOperationException($"A preset called '{name}' already exists.");
        }

        ini.AddPreset(name);
        ini.Set("VideoCopies", "1", name);
        ini.Set("ImageCopies", "1", name);
        ini.Save(Paths.ConfigFile);

        Load();
        Selected = Presets.FirstOrDefault(p => p.Name == name);
        IsDirty = true;
    }

    /// <summary>
    /// Removes a preset from the config. Its folders and any media in them are left alone: a
    /// settings edit should never delete someone's files.
    /// </summary>
    public void RemovePreset(string name)
    {
        var ini = IniFile.Load(Paths.ConfigFile);
        ini.RemovePreset(name);
        ini.Save(Paths.ConfigFile);

        Load();
        IsDirty = true;
    }

    public static string DescribeInts(string video, string image)
    {
        static int Parse(string value) =>
            int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : 0;

        var v = Parse(video);
        var i = Parse(image);

        return (v, i) switch
        {
            (0, 0) => "produces nothing",
            (0, _) => $"{i} per image, video ignored",
            (_, 0) => $"{v} per video, images ignored",
            _ when v == i => $"{v} per file",
            _ => $"{v} per video, {i} per image",
        };
    }
}
