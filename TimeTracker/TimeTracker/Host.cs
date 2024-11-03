using Microsoft.Extensions.DependencyInjection;
using TimeTracker.ViewModels;
using TimeTracker.Views;

namespace TimeTracker
{
    /// <summary>
    ///     Provides a host for the application's services and manages their lifetimes
    /// </summary>
    public static class Host
    {
        private static IServiceProvider _serviceProvider;

        /// <summary>
        ///     Starts the host and configures the application's services
        /// </summary>
        public static void Start()
        {
            var services = new ServiceCollection();

            services.AddTransient<TimeTrackerViewModel>();
            services.AddTransient<TimeTrackerView>();

            _serviceProvider = services.BuildServiceProvider();
        }

        /// <summary>
        ///     Get service of type <typeparamref name="T"/>
        /// </summary>
        /// <typeparam name="T">The type of service object to get</typeparam>
        /// <exception cref="System.InvalidOperationException">There is no service of type <typeparamref name="T"/></exception>
        public static T GetService<T>() where T : class
        {
            return _serviceProvider.GetRequiredService<T>();
        }
    }
}