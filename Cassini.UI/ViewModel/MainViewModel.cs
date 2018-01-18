using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using Cassini.DA;
using Cassini.Model;
using Cassini.UI.Event;
using Cassini.UI.Service;
using Prism.Events;
using Prism.Mvvm;

namespace Cassini.UI.ViewModel
{
    public class MainViewModel : BindableBase
    {
        public MainViewModel(
            IDirectionsViewModel directionsViewModel,
            IParametersViewModel parametersViewModel,
            IAgentActsViewModel agentActsViewModel
            )
        {
            DirectionsViewModel = directionsViewModel;
            ParametersViewModel = parametersViewModel;
            AgentActsViewModel = agentActsViewModel;
        }
        

        public IDirectionsViewModel DirectionsViewModel { get; }
        public IParametersViewModel ParametersViewModel { get; }
        public IAgentActsViewModel AgentActsViewModel { get; }

        public async Task LoadAsync()
        {
            await DirectionsViewModel.LoadAsync();
            await ParametersViewModel.GetCommissionTypesAsync();
            await ParametersViewModel.GetActStatusAsync();
            await ParametersViewModel.GetAgentChanelsAsync();
        }



        
    }
}