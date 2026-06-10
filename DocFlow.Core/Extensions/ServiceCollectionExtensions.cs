using System;
using System.Collections.Generic;
using DocFlow.Core.Factory;
using DocFlow.Core.Helpers;
using DocFlow.Core.Interfaces;
using DocFlow.Core.Models;
using DocFlow.Core.Services;

namespace DocFlow.Core.Extensions
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddDocFlowCore(this IServiceCollection services)
        {
            return AddDocFlowCore(services, null, null);
        }

        public static IServiceCollection AddDocFlowCore(this IServiceCollection services, ILogger logger)
        {
            return AddDocFlowCore(services, logger, null);
        }

        public static IServiceCollection AddDocFlowCore(this IServiceCollection services, ILogger logger, DocFlowSettings settings)
        {
            if (services == null)
            {
                throw new ArgumentNullException(nameof(services));
            }

            var effectiveLogger = logger ?? new NullLogger();
            var effectiveSettings = settings ?? DocFlowSettings.CreateDefault();
            services.AddSingleton<ILogger>(provider => effectiveLogger);
            services.AddSingleton<DocFlowSettings>(provider => effectiveSettings);
            services.AddSingleton<DocumentFactory>(provider => new DocumentFactory());
            services.AddSingleton<IWordService>(provider => new WordService(provider.GetRequiredService<ILogger>(), provider.GetRequiredService<DocFlowSettings>()));
            services.AddSingleton<IExcelService>(provider => new ExcelService(provider.GetRequiredService<ILogger>(), provider.GetRequiredService<DocFlowSettings>()));
            services.AddSingleton<ICsvService>(provider => new CsvService(provider.GetRequiredService<IExcelService>(), provider.GetRequiredService<ILogger>(), provider.GetRequiredService<DocFlowSettings>()));
            services.AddSingleton<IHtmlService>(provider => new HtmlService(provider.GetRequiredService<IWordService>(), provider.GetRequiredService<IExcelService>(), provider.GetRequiredService<ILogger>(), provider.GetRequiredService<DocFlowSettings>()));
            services.AddSingleton<IImageService>(provider => new ImageService(provider.GetRequiredService<IWordService>(), provider.GetRequiredService<IExcelService>(), provider.GetRequiredService<ILogger>(), provider.GetRequiredService<DocFlowSettings>()));
            services.AddSingleton<IPdfService>(provider =>
                new PdfService(
                    provider.GetRequiredService<IWordService>(),
                    provider.GetRequiredService<IExcelService>(),
                    provider.GetRequiredService<ILogger>(),
                    provider.GetRequiredService<DocFlowSettings>()));
            services.AddSingleton<IConversionService>(provider =>
                new ConversionService(
                    provider.GetRequiredService<IWordService>(),
                    provider.GetRequiredService<IExcelService>(),
                    provider.GetRequiredService<IPdfService>(),
                    provider.GetRequiredService<ICsvService>(),
                    provider.GetRequiredService<IHtmlService>(),
                    provider.GetRequiredService<IImageService>(),
                    provider.GetRequiredService<ILogger>(),
                    provider.GetRequiredService<DocFlowSettings>()));

            return services;
        }
    }

    public interface IServiceCollection
    {
        IServiceCollection AddSingleton<TService>(Func<IServiceCollection, TService> factory) where TService : class;

        TService GetRequiredService<TService>() where TService : class;
    }

    public sealed class ServiceCollection : IServiceCollection
    {
        private readonly IDictionary<Type, object> _instances = new Dictionary<Type, object>();
        private readonly object _syncRoot = new object();

        public IServiceCollection AddSingleton<TService>(Func<IServiceCollection, TService> factory) where TService : class
        {
            if (factory == null)
            {
                throw new ArgumentNullException(nameof(factory));
            }

            lock (_syncRoot)
            {
                _instances[typeof(TService)] = new Lazy<object>(() => factory(this), true);
            }

            return this;
        }

        public TService GetRequiredService<TService>() where TService : class
        {
            lock (_syncRoot)
            {
                object instance;
                if (!_instances.TryGetValue(typeof(TService), out instance))
                {
                    throw new InvalidOperationException("The requested service is not registered: " + typeof(TService).FullName);
                }

                var lazy = instance as Lazy<object>;
                return (TService)(lazy != null ? lazy.Value : instance);
            }
        }
    }
}
