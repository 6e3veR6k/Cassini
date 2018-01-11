using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using Cassini.Model;
using Cassini.UI.Service;

namespace Cassini.UI.ViewModel
{
    public class MainViewModel : BaseView
    {

        public MainViewModel(
            IDirectionsViewModel directionsViewModel,
            IParametersViewModel parametersViewModel)
        {
            DirectionsViewModel = directionsViewModel;
            ParametersViewModel = parametersViewModel;
        }

        public IDirectionsViewModel DirectionsViewModel { get; }
        public IParametersViewModel ParametersViewModel { get; }

        public async Task LoadAsync()
        {
            await DirectionsViewModel.LoadAsync();
            await ParametersViewModel.GetCommissionTypesAsync();
            await ParametersViewModel.GetActStatusAsync();
            await ParametersViewModel.GetAgentChanelsAsync();

        }
    }
}