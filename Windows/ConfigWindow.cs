using Dalamud.Interface.Windowing;
using Dalamud.Bindings.ImGui;
using EZSheets.Services;

namespace EZSheets.Windows;

public sealed class ConfigWindow : Window, IDisposable
{
    private readonly Configuration configuration;
    private readonly SupabaseRestClient supabase;

    public ConfigWindow(Configuration configuration, SupabaseRestClient supabase)
        : base("EZSheets Settings###EZSheetsConfigWindow", ImGuiWindowFlags.AlwaysAutoResize)
    {
        this.SizeCondition = ImGuiCond.FirstUseEver;
        this.configuration = configuration;
        this.supabase = supabase;
    }

    public void Dispose()
    {
    }

    public override void Draw()
    {
        ImGui.TextWrapped("EZSheets now uses direct Discord OAuth for login and uses Supabase only as the cloud backend.");
        ImGui.Separator();

        ImGui.TextUnformatted("Supabase Project URL:");
        ImGui.InputText("##ProjectUrl", ref UnsafeProjectUrlRef.Value, 512, ImGuiInputTextFlags.ReadOnly);
        ImGui.TextUnformatted("Publishable Key:");
        ImGui.InputText("##PublishableKey", ref UnsafePublishableKeyRef.Value, 1024, ImGuiInputTextFlags.ReadOnly);
        ImGui.TextUnformatted("Discord Client ID:");
        ImGui.InputText("##DiscordClientId", ref UnsafeDiscordClientIdRef.Value, 64, ImGuiInputTextFlags.ReadOnly);
        ImGui.TextUnformatted("Desktop OAuth Redirect URL:");
        ImGui.InputText("##OAuthRedirect", ref UnsafeRedirectUrlRef.Value, 256, ImGuiInputTextFlags.ReadOnly);

        ImGui.Spacing();
        ImGui.TextWrapped("Setup checklist:");
        ImGui.BulletText("In the Discord Developer Portal, add the redirect URL above to your application Redirects.");
        ImGui.BulletText("In Supabase, deploy the Edge Function from supabase/functions/EZSheets-api with --no-verify-jwt.");
        ImGui.BulletText("Set the Edge Function secrets DISCORD_CLIENT_ID and DISCORD_CLIENT_SECRET in Supabase.");
        ImGui.BulletText("Run the SQL from supabase/setup.sql to create the new Discord-based EZSheets tables.");
        ImGui.BulletText("You do not need to configure a Supabase Auth Discord provider for this build.");

        ImGui.Spacing();
        ImGui.TextUnformatted($"Configured Project: {(this.supabase.HasConfiguredProject ? "Yes" : "No")}");
        ImGui.TextUnformatted($"Signed In: {(this.supabase.IsAuthenticated ? "Yes" : "No")}");
    }

    private static class UnsafeProjectUrlRef
    {
        public static string Value = Configuration.EmbeddedSupabaseUrl;
    }

    private static class UnsafePublishableKeyRef
    {
        public static string Value = Configuration.EmbeddedSupabasePublishableKey;
    }

    private static class UnsafeDiscordClientIdRef
    {
        public static string Value = Configuration.EmbeddedDiscordClientId;
    }

    private static class UnsafeRedirectUrlRef
    {
        public static string Value = Configuration.DiscordCallbackUrl;
    }
}
