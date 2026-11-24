using Hangfire;
using Hangfire.Console;
using Hangfire.PostgreSql;
using Npgsql;
using RfpDashboard.Jobs;

var builder = WebApplication.CreateBuilder(args);

var connectionString = builder.Configuration.GetConnectionString("DefaultConnection")
    ?? throw new InvalidOperationException("Connection string 'DefaultConnection' is not configured.");

builder.Services.AddScoped(_ => new NpgsqlConnection(connectionString));
builder.Services.AddScoped<CalProcureSyncJob>();
builder.Services.AddScoped<CalProcureClassifyJob>();
builder.Services.AddHttpClient();

builder.Services.AddHangfire(configuration =>
    configuration
        .UseConsole()
        .UsePostgreSqlStorage(options =>
            options.UseNpgsqlConnection(connectionString)));

builder.Services.AddHangfireServer();

var app = builder.Build();

app.UseHangfireDashboard("/hangfire");

RecurringJob.AddOrUpdate<CalProcureSyncJob>(
    "cal-procure-sync",
    job => job.ExecuteAsync(null),
    Cron.Never);

app.MapGet("/", () => "Hello World!");

app.Run();
