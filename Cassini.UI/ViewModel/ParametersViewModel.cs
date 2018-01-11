using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using Cassini.Model;
using Cassini.UI.Service;

namespace Cassini.UI.ViewModel
{
    public class ParametersViewModel : BaseView, IParametersViewModel
    {
        private IActsParametersDataService _actsParametersDataService;


        private DateTime _periodDate;
        private DateTime _startDate;
        private DateTime _endDate;


        

        public ObservableCollection<ActStatus> ActStatuses { get; }
        public ObservableCollection<Commission> CommissionTypes { get; }
        public ObservableCollection<Chanel> Chanels { get; }

        public ParametersViewModel(IActsParametersDataService actsParametersDataService)
        {
            _actsParametersDataService = actsParametersDataService;
            ActStatuses = new ObservableCollection<ActStatus>();
            CommissionTypes = new ObservableCollection<Commission>();
            Chanels = new ObservableCollection<Chanel>();
        }

        public async Task GetActStatusAsync()
        {
            var statuses = await _actsParametersDataService.GetActStatusAsync();
            ActStatuses.Clear();

            foreach (var status in statuses)
            {
                ActStatuses.Add(status);
            }
        }

        public async Task GetAgentChanelsAsync()
        {
            var chanels = await _actsParametersDataService.GetAgentChanelsAsync();
            Chanels.Clear();

            foreach (var chanel in chanels)
            {
                Chanels.Add(chanel);
            }
        }

        public async Task GetCommissionTypesAsync()
        {
            var types = await _actsParametersDataService.GetCommissionTypesAsync();
            CommissionTypes.Clear();

            foreach (var cType in types)
            {
                CommissionTypes.Add(cType);
            }
        }

    }
}