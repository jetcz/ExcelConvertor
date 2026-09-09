using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

using ExcelConvertor.Lib;
using ExcelConvertor.WebApp;

var builder = WebApplication.CreateBuilder();
const long maxUploadSizeBytes = 10 * 1024 * 1024;

builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = maxUploadSizeBytes;
});

builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = maxUploadSizeBytes;
});

builder.Services.AddAuthentication();
builder.Services.AddTransient<IExcelToJsonConverter, ExcelToJsonConverter>();
builder.Services.AddDotVVM<DotvvmStartup>();

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();
app.UseAuthorization();

app.UseDotVVM<DotvvmStartup>();
app.MapDotvvmHotReload();

app.Run();