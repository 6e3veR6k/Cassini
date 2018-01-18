using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Cassini.Model;
using Cassini.UI.Event;
using Cassini.UI.Service;
using Prism.Events;
using Prism.Mvvm;

namespace Cassini.UI.ViewModel
{
    public class AgentActsViewModel : BindableBase, IAgentActsViewModel
    {
        private IAgetActsCommissionDataService _agetActsCommissionDataService;
        private ObservableCollection<ActsResultSet> _actsResultSet;
        private ObservableCollection<ActResultSetSumView> _actsResultSetSum;
        private IEventAggregator _eventAggregator;

        public ObservableCollection<ActsResultSet> ActsResultSet
        {
            get { return _actsResultSet; }
            set { _actsResultSet = value; }
        }

        public ObservableCollection<ActResultSetSumView> ActsResultSetSum
        {
            get { return _actsResultSetSum; }
            set { _actsResultSetSum = value; }
        }

        public AgentActsViewModel(IEventAggregator eventAggregator,
            IAgetActsCommissionDataService agetActsCommissionDataService)
        {
            _agetActsCommissionDataService = agetActsCommissionDataService;
            _eventAggregator = eventAggregator;
            _eventAggregator.GetEvent<OnParametersButtonClickEvent>().Subscribe(Action);

            ActsResultSet = new ObservableCollection<ActsResultSet>();
            ActsResultSetSum = new ObservableCollection<ActResultSetSumView>();
        }


        private async void Action(InputParametersModel inputParametersModel)
        {
            ActsResultSetSum.Clear();
            ActsResultSet.Clear();
            await GetCommissionActs(inputParametersModel);
        }

        public async Task GetCommissionActs(InputParametersModel inputParameters)
        {

            var result = await _agetActsCommissionDataService.GetActsComissionsSumResults(inputParameters);
            var actsResultSets = result as IList<ActsResultSet> ?? result.ToList();

            LoadActsResultSetSum(actsResultSets);
            _eventAggregator.GetEvent<ActResultSetLoadEvent>().Publish(actsResultSets);
        }

        private void LoadActsResultSetSum(IList<ActsResultSet> actsResultSets)
        {
            var resultSum = from r in actsResultSets
                group r by r.ActId
                into g
                select new ActResultSetSumView
                {
                    ActId = g.Key,
                    SummCommission = g.Sum(r => r.CommissionValue),
                    AgentName = g.Max(r => r.AgentName),
                    DirCode = g.Max(r => r.BranchCode.Substring(0, 2)),
                    DogType = g.Max(r => r.DocumentType),
                    INN = g.Max(r => r.IdentificationCodeEDRPOU)
                };

            foreach (var actsResultSet in resultSum)
            {
                ActsResultSetSum.Add(actsResultSet);
            }
        }

        private void LoadResultDataSet(IList<ActsResultSet> actsResultSets)
        {
            foreach (var actsResultSet in actsResultSets)
            {
                ActsResultSet.Add(actsResultSet);
            }
        }

    }
}