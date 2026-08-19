using GeneradorFIT.ViewModels;
using GeneradorFIT.Views;
using Microsoft.Extensions.DependencyInjection;
using System.Windows;

namespace GeneradorFIT
{
    public partial class App : Application
    {
        private ServiceProvider _serviceProvider;

        public App()
        {
            var services = new ServiceCollection();
            ConfigureServices(services);
            _serviceProvider = services.BuildServiceProvider();
        }

        private void ConfigureServices(ServiceCollection services)
        {
            services.AddTransient<GeneradorFitViewModel>();

            services.AddSingleton<MainWindow>(provider =>
            {
                var window = new MainWindow
                {
                    DataContext = provider.GetRequiredService<GeneradorFitViewModel>()
                };
                return window;
            });
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
            mainWindow.Show();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _serviceProvider?.Dispose();
            base.OnExit(e);
        }
    }
}
