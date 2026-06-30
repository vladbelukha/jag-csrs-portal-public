using Csrs.Api;
using Csrs.Api.Authentication;
using Csrs.Api.Configuration;
using Csrs.Api.Models;
using Csrs.Api.Services;
using Csrs.Api.ApiGateway;
using System.Configuration;
using Grpc.Net.Client;
using Grpc.Core;
using Grpc.Net.Client.Configuration;
using Serilog;
using Csrs.Interfaces.Dynamics;
using System.Net;

namespace Microsoft.Extensions.DependencyInjection;

public static class WebApplicationBuilderExtensions
{
    /// <summary>
    /// Adds repository and dependant services.
    /// </summary>
    /// <param name="builder"></param>
    public static void AddServices(this WebApplicationBuilder builder)
    {
        ArgumentNullException.ThrowIfNull(builder);

        var logger = Log.ForContext(typeof(WebApplicationBuilderExtensions));

        var configuration = builder.Configuration.Get<CsrsConfiguration>();
        OAuthConfiguration? oAuthOptions = configuration?.OAuth;

        if (string.IsNullOrEmpty(oAuthOptions?.ResourceUrl))
        {
            const string message = "OAuth configuration is not set";
            logger.Error(message);
            throw new ConfigurationErrorsException(message);
        }

        ApiGatewayOptions? apiGatewayOptions = configuration?.ApiGateway;
        if (string.IsNullOrEmpty(apiGatewayOptions?.BasePath))
        {
            const string message = "ApiGateWay configuration is not set";
            logger.Error(message);
            throw new ConfigurationErrorsException(message);
        }

        var services = builder.Services;

        logger.Debug("Setting up oAuthOptions and apiGatewayOptions");
        services.AddSingleton(oAuthOptions);
        services.AddSingleton(apiGatewayOptions);

        logger.Debug("Adding memory cache");
        services.AddMemoryCache();

        // Add OAuth Middleware
        services.AddTransient<OAuthHandler>();
        // Add ApiGateway Middleware
        services.AddTransient<ApiGatewayHandler>();

        // Register IOAuthApiClient
        services.AddHttpClient<IOAuthApiClient, OAuthApiClient>(client =>
        {
            client.Timeout = TimeSpan.FromSeconds(15); // set the auth timeout
        });
        services.AddSingleton(new DynamicsClientOptions { NativeOdataResourceUrl = oAuthOptions.ResourceUrl });
        services.AddHttpClient<IDynamicsClient, DynamicsClient>(client =>
        {

            client.BaseAddress = new Uri(apiGatewayOptions.BasePath);
            client.Timeout = TimeSpan.FromSeconds(30); // data timeout
            //client.BaseAddress = new Uri(oAuthOptions.ResourceUrl);
            //client.Timeout = TimeSpan.FromSeconds(300); // data timeout

        })
        .AddHttpMessageHandler<OAuthHandler>()
        .AddHttpMessageHandler<ApiGatewayHandler>();

        logger.Debug("Configuing FileManager Service");
        ConfigureFileManagerService(builder, configuration?.FileManager);

        services.AddHttpContextAccessor();

        // Add services
        services.AddTransient<ITokenService, TokenService>();
        services.AddTransient<IMessageService, MessageService>();
        services.AddTransient<IAccountService, AccountService>();
        services.AddTransient<IFileService, FileService>();
        services.AddTransient<IUserService, UserService>();
        services.AddTransient<ILookupService, LookupService>();
        services.AddTransient<IDocumentService, DocumentService>();
        services.AddTransient<ITaskService, TaskService>();

    }

    private static void ConfigureFileManagerService(WebApplicationBuilder builder, FileManagerConfiguration? configuration)
    {
        var logger = Log.ForContext(typeof(WebApplicationBuilderExtensions));

        if (string.IsNullOrWhiteSpace(configuration?.Address))
        {
            const string message = $"FileManager configuration is not set, {nameof(CsrsConfiguration.FileManager)}:{nameof(FileManagerConfiguration.Address)} is required.";
            logger.Error(message);
            throw new ConfigurationErrorsException(message);
        }

        string address = configuration.Address;

        // determine if we are using http or https
        ChannelCredentials credentials;

        if (configuration.Secure.HasValue && !configuration.Secure.Value)
        {
            logger.Information("Configuration explicitly set Secure=false. Using insecure channel for File Manager service.");
            // Required for Grpc.Net.Client to use h2c (unencrypted HTTP/2) on .NET 6
            AppContext.SetSwitch("System.Net.Http.SocketsHttpHandler.Http2UnencryptedSupport", true);
            credentials = ChannelCredentials.Insecure;
        }
        else
        {
            logger.Information("Configuration Secure=true or not set. Using secure channel for File Manager service.");
            credentials = ChannelCredentials.SecureSsl;
        }

        //credentials = ChannelCredentials.SecureSsl;
        logger.Information("Using file manager service {Address}", address);
        builder.Services.AddSingleton(services =>
        {
            var channel = GrpcChannel.ForAddress(address, new GrpcChannelOptions
            {
                Credentials = credentials,
                ServiceConfig = new ServiceConfig { LoadBalancingConfigs = { new RoundRobinConfig() } },
                ServiceProvider = services

            });

            return channel;
        });

        builder.Services.AddTransient(services =>
        {
            GrpcChannel channel = services.GetRequiredService<GrpcChannel>();
            return new Csrs.Services.FileManager.FileManager.FileManagerClient(channel);
        });
    }

    /// <summary>
    /// Gets a logger for application setup.
    /// </summary>
    private static Serilog.ILogger GetLogger()
    {
        var logger = new LoggerConfiguration()
            .WriteTo.Console()
            .WriteTo.Debug()
            .CreateLogger();

        return logger;
    }
}
