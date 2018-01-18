using Autofac;
using Cassini.DA;
using Cassini.UI.Service;
using Cassini.UI.ViewModel;
using Prism.Events;

namespace Cassini.UI.StartUp
{
    public class Bootstrapper
    {
        public IContainer Bootstrap()
        {
            var builder = new ContainerBuilder();

            builder.RegisterType<CallistoDb>().AsSelf();

            builder.RegisterType<EventAggregator>().As<IEventAggregator>().SingleInstance();

            builder.RegisterType<DirectionsViewModel>().As<IDirectionsViewModel>();
            builder.RegisterType<ParametersViewModel>().As<IParametersViewModel>();
            builder.RegisterType<AgentActsViewModel>().As<IAgentActsViewModel>();

            builder.RegisterType<AgetActsCommissionDataService>().As<IAgetActsCommissionDataService>();
            builder.RegisterType<DirectionDataSevice>().As<IDirectionDataSevice>();
            builder.RegisterType<ActsParametersDataService>().As<IActsParametersDataService>();

            builder.RegisterType<MainViewModel>().AsSelf();
            builder.RegisterType<MainWindow>().AsSelf();

            return builder.Build();
        }
    }
}