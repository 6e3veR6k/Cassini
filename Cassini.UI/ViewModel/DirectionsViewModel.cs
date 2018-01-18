using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Cassini.Model;
using Cassini.UI.Event;
using Cassini.UI.Service;
using Prism.Commands;
using Prism.Events;
using Prism.Mvvm;

namespace Cassini.UI.ViewModel
{
    public class DirectionsViewModel : BindableBase, IDirectionsViewModel
    {
        private IDirectionDataSevice _directionDataService;
        private ObservableCollection<DirectionViewModel> _directions;
        private IEventAggregator _eventAggregator;

        public DirectionsViewModel(IDirectionDataSevice directionDataSevice, IEventAggregator eventAggregator)
        {
            _directionDataService = directionDataSevice;

            Directions = new ObservableCollection<DirectionViewModel>();
            _eventAggregator = eventAggregator;

            
        }

        public ObservableCollection<DirectionViewModel> Directions
        {
            get => _directions;
            set
            {
                _directions = value;
            }
        }

        public ObservableCollection<DirectionViewModel> SelectedDirections
        {
            get
            {
                return new ObservableCollection<DirectionViewModel>(Directions.Where(d => d.IsSelected));
            }
        }
        
        public async Task LoadAsync()
        {
            var directions = await _directionDataService.GetAllAsync();

            Directions.Clear();

            foreach (var direction in directions)
            {
                Directions.Add(new DirectionViewModel
                {
                    Code = direction.Code,
                    Guid = direction.Guid,
                    Title = direction.Title,
                    IsSelected = false
                });
            }

            foreach (var dir in Directions)
            {
                dir.PropertyChanged += this.OnDirectionViewModelPropertyChanged;
            }
        }

        private void OnDirectionViewModelPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            string IsSelected = "IsSelected";

            if (e.PropertyName == IsSelected)
                this.RaisePropertyChanged(nameof(SelectedDirections));

            _eventAggregator.GetEvent<SelectedDirectionsEvent>().Publish(SelectedDirections);
        }
    }
}