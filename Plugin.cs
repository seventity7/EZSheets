using Dalamud.Game.Command;
using Dalamud.Game.Gui.Toast;
using Dalamud.Interface.ImGuiNotification;
using Dalamud.Interface.Windowing;
using Dalamud.IoC;
using Dalamud.Plugin;
using Dalamud.Plugin.Services;
using EZSheets.Services;
using EZSheets.Windows;

namespace EZSheets;

public sealed class Plugin : IDalamudPlugin
{
    public const string CommandName = "/EZSheets";
    public const string CommandAliasShort = "/ss";
    public const string CommandAliasSheet = "/sheet";

    private readonly WindowSystem windowSystem = new("EZSheets");
    private readonly CommandInfo commandInfo;

    [PluginService]
    internal static IDalamudPluginInterface PluginInterface { get; private set; } = null!;

    [PluginService]
    internal static ICommandManager CommandManager { get; private set; } = null!;

    [PluginService]
    internal static IPluginLog Log { get; private set; } = null!;

    [PluginService]
    internal static IChatGui ChatGui { get; private set; } = null!;

    [PluginService]
    internal static IObjectTable ObjectTable { get; private set; } = null!;

    [PluginService]
    internal static IToastGui ToastGui { get; private set; } = null!;

    [PluginService]
    internal static INotificationManager NotificationManager { get; private set; } = null!;

    public Plugin()
    {
        this.commandInfo = new CommandInfo(this.OnCommand)
        {
            HelpMessage = "Open the EZSheets collaborative spreadsheet window.",
        };

        this.Configuration = PluginInterface.GetPluginConfig() as Configuration ?? new Configuration();
        this.Configuration.EnsureEmbeddedDefaults();
        this.Configuration.Save();

        this.Supabase = new SupabaseRestClient(this.Configuration);
        this.MainWindow = new MainWindow(this.Configuration, this.Supabase, () => { });

        this.windowSystem.AddWindow(this.MainWindow);

        CommandManager.AddHandler(CommandName, this.commandInfo);
        CommandManager.AddHandler(CommandAliasShort, this.commandInfo);
        CommandManager.AddHandler(CommandAliasSheet, this.commandInfo);

        PluginInterface.UiBuilder.Draw += this.DrawUi;
        PluginInterface.UiBuilder.OpenMainUi += this.OpenMainUi;

        _ = this.MainWindow.TryRestoreSessionAndLoadAsync();
    }

    public string Name => "EZSheets";

    public Configuration Configuration { get; }

    public SupabaseRestClient Supabase { get; }

    public MainWindow MainWindow { get; }

    public void Dispose()
    {
        this.windowSystem.RemoveAllWindows();
        this.MainWindow.Dispose();
        this.Supabase.Dispose();

        PluginInterface.UiBuilder.Draw -= this.DrawUi;
        PluginInterface.UiBuilder.OpenMainUi -= this.OpenMainUi;

        CommandManager.RemoveHandler(CommandName);
        CommandManager.RemoveHandler(CommandAliasShort);
        CommandManager.RemoveHandler(CommandAliasSheet);
    }

    private void OnCommand(string command, string arguments)
    {
        _ = command;
        _ = arguments;
        this.MainWindow.IsOpen = true;
    }

    private void DrawUi()
    {
        this.windowSystem.Draw();
    }

    private void OpenMainUi()
    {
        this.MainWindow.IsOpen = true;
    }
}
