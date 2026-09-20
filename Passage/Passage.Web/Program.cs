using Passage.Core.Extensibility;
using Passage.Web.Components;
using Passage.Web.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.AddSingleton<ScriptLibrary>();
builder.Services.AddSingleton<ExportService>();
builder.Services.AddSingleton<GoalSettingsStore>();
builder.Services.AddSingleton<SlateStore>();
// The assist seam (docs/SLATE-PLAN.md, Phase 7). Only the null partner
// exists; it asks nothing and sends nothing anywhere. A real partner is
// registered here in its place and nowhere else — the dock shows its
// buttons only when the registered partner is not the null one.
builder.Services.AddSingleton<IStoryPartner, NullStoryPartner>();

var app = builder.Build();

app.MapStaticAssets();
app.UseAntiforgery();

// Export the posted script text in the requested format and stream it back
// as a download. The content is posted (rather than read from the library)
// so unsaved work can be exported too.
app.MapPost("/api/export", async (HttpContext context, ExportService exports) =>
{
    var form = await context.Request.ReadFormAsync();
    var content = form["content"].ToString();
    var format = form["format"].ToString();
    var name = form["name"].ToString();

    var result = exports.Export(content, format, name);
    if (result is null)
    {
        return Results.BadRequest($"Unknown export format '{format}'.");
    }

    return Results.File(result.Value.Bytes, result.Value.ContentType, result.Value.FileName);
}).DisableAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
