using System.Collections.ObjectModel;
using System.Threading.Tasks;
using Cassini.UI.Service;

namespace Cassini.UI.ViewModel
{
    public class DirectionsViewModel : BaseView, IDirectionsViewModel
    {
        private IDirectionDataSevice _directionDataService;

        public DirectionsViewModel(IDirectionDataSevice directionDataSevice)
        {
            _directionDataService = directionDataSevice;
            Directions = new ObservableCollection<DirectionViewModel>();
        }

        public ObservableCollection<DirectionViewModel> Directions { get; }

        public async Task LoadAsync()
        {
            var directions = await _directionDataService.GetAllAsync();

            Directions.Clear();

            foreach (var direction in directions)
            {
                Directions.Add(new DirectionViewModel(direction));
            }
        }
    }
}