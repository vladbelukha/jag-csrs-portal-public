using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;

namespace Microsoft.Extensions.DependencyInjection;

public static class AuthenticationExtensions
{
    /// <summary>
    /// Authentication scheme name used by Dynamics-to-API requests (symmetric key JWT).
    /// </summary>
    public const string DynamicsAuthenticationScheme = "DynamicsBearer";

    /// <summary>
    /// Authorization policy name that requires the Dynamics authentication scheme.
    /// </summary>
    public const string DynamicsPolicy = "DynamicsPolicy";

    /// <summary>
    /// Internal routing scheme that forwards each request to exactly one JWT handler,
    /// preventing the OIDC handler from attempting to validate Dynamics tokens and vice versa.
    /// </summary>
    private const string RouterScheme = "SmartBearer";

    public static void AddJwtBearerAuthentication(this WebApplicationBuilder builder)
    {
        builder.Services
            .AddAuthentication(options =>
            {
                // The router scheme is the single default. It inspects the request path and
                // delegates to the appropriate JWT Bearer scheme before any token validation occurs.
                // This prevents the OIDC handler from seeing Dynamics tokens (and logging false
                // validation failures) and keeps both authentication flows completely isolated.
                options.DefaultScheme = RouterScheme;
                options.DefaultChallengeScheme = RouterScheme;
                options.DefaultForbidScheme = RouterScheme;
            })
            .AddPolicyScheme(RouterScheme, "Route to FE or Dynamics JWT bearer", opts =>
            {
                opts.ForwardDefaultSelector = context =>
                {
                    // Dynamics service-to-service calls are served exclusively under /api/dynamics.
                    if (context.Request.Path.StartsWithSegments("/api/dynamics", StringComparison.OrdinalIgnoreCase))
                    {
                        return DynamicsAuthenticationScheme;
                    }

                    // All portal (front-end user) requests use the OIDC-backed Bearer scheme.
                    return JwtBearerDefaults.AuthenticationScheme;
                };
            })
            // ── Scheme 1: Front-end / portal users ──────────────────────────────────────
            // OIDC JWT issued by the identity provider after interactive login.
            // https://github.com/dotnet/aspnetcore/blob/v6.0.2/src/Security/Authentication/JwtBearer/samples/JwtBearerSample/Startup.cs
            .AddJwtBearer(JwtBearerDefaults.AuthenticationScheme, options =>
            {
                options.Authority = builder.Configuration["Jwt:Authority"];
                options.Audience = builder.Configuration["Jwt:Audience"];
                options.Events = new JwtBearerEvents
                {
                    OnTokenValidated = ctx =>
                    {
                        var log = ctx.HttpContext.RequestServices.GetRequiredService<ILoggerFactory>()
                            .CreateLogger("Auth.Bearer");
                        log.LogDebug("[Bearer/OIDC] Token validated. Subject={Subject}",
                            ctx.Principal?.FindFirst("sub")?.Value);
                        return Task.CompletedTask;
                    },
                    OnAuthenticationFailed = ctx =>
                    {
                        var log = ctx.HttpContext.RequestServices.GetRequiredService<ILoggerFactory>()
                            .CreateLogger("Auth.Bearer");
                        log.LogWarning(ctx.Exception,
                            "[Bearer/OIDC] Token validation failed. Path={Path} Exception={Message}",
                            ctx.Request.Path, ctx.Exception.Message);
                        return Task.CompletedTask;
                    }
                };
            })
            // ── Scheme 2: Dynamics service account ──────────────────────────────────────
            // Symmetric key JWT for machine-to-machine calls originating from Dynamics.
            // Configure JWT:Dynamics:Secret, JWT:Dynamics:Issuer, JWT:Dynamics:Audience
            // in secrets.json / environment variables.
            .AddJwtBearer(DynamicsAuthenticationScheme, options =>
            {
                options.Authority = builder.Configuration["Jwt:Dynamics:Authority"];
                options.Audience = builder.Configuration["Jwt:Dynamics:Audience"];
                options.Events = new JwtBearerEvents
                {
                    OnTokenValidated = ctx =>
                    {
                        var log = ctx.HttpContext.RequestServices.GetRequiredService<ILoggerFactory>()
                            .CreateLogger("Auth.DynamicsBearer");
                        log.LogDebug("[DynamicsBearer] Token validated. Subject={Subject}",
                            ctx.Principal?.FindFirst("sub")?.Value);
                        return Task.CompletedTask;
                    },
                    OnAuthenticationFailed = ctx =>
                    {
                        var log = ctx.HttpContext.RequestServices.GetRequiredService<ILoggerFactory>()
                            .CreateLogger("Auth.DynamicsBearer");
                        log.LogWarning(ctx.Exception,
                            "[DynamicsBearer] Token validation failed. Path={Path} KeyConfigured={KeyPresent} Exception={Message}",
                            ctx.Request.Path,
                            !string.IsNullOrEmpty(builder.Configuration["JWT:Dynamics:Secret"]),
                            ctx.Exception.Message);
                        return Task.CompletedTask;
                    }
                };
            });

        builder.Services.AddAuthorization(options =>
        {
            options.AddPolicy(DynamicsPolicy, policy =>
            {
                policy.AuthenticationSchemes.Add(DynamicsAuthenticationScheme);
                policy.RequireAuthenticatedUser();
            });
        });
    }
}
