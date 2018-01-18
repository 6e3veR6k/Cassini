using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Cassini.Model;
using Cassini.UI.Event;
using Cassini.UI.Service;
using Microsoft.Win32;
using Prism.Commands;
using Prism.Events;
using Prism.Mvvm;

namespace Cassini.UI.ViewModel
{
    public class ParametersViewModel : BindableBase, IParametersViewModel, INotifyDataErrorInfo
    {
        private IActsParametersDataService _actsParametersDataService;

        private DateTime? _periodDate;
        private DateTime? _startDate;
        private ActStatus _selectedActStatus;
        private Commission _selectedCommission;
        private IEventAggregator _eventAggregator;
        private InputParametersModel _inputParametersModel;
        private string _resultTextSet;


        public ParametersViewModel(IActsParametersDataService actsParametersDataService, IEventAggregator eventAggregator)
        {
            _actsParametersDataService = actsParametersDataService;
            ActStatuses = new ObservableCollection<ActStatus>();
            CommissionTypes = new ObservableCollection<Commission>();
            Chanels = new ObservableCollection<ChanelViewModel>();
            SelectedDirections = new ObservableCollection<DirectionViewModel>();

            _eventAggregator = eventAggregator;
            _eventAggregator.GetEvent<SelectedDirectionsEvent>().Subscribe(GetDirections);
            _eventAggregator.GetEvent<ActResultSetLoadEvent>().Subscribe(LoadResultSet);


            OnViewReportButtonClick = new DelegateCommand(OnClickedViewReportButton, CanViewReport);
            OnExportDataButtonClick = new DelegateCommand(OnClickedExportDataButton, CanClickExport);
        }


        public ICommand OnViewReportButtonClick { get; set; }
        public ICommand OnExportDataButtonClick { get; set; }
        

        public DateTime? PeriodDateTime
        {
            get => _periodDate;
            set
            {
                _periodDate = value; 
                RaisePropertyChanged(nameof(PeriodDateTime));
                ((DelegateCommand)OnViewReportButtonClick).RaiseCanExecuteChanged();
            }
        }

        public DateTime? StartDateTime
        {
            get => _startDate;
            set
            {
                _startDate = value;
                RaisePropertyChanged(nameof(StartDateTime));
                ((DelegateCommand)OnViewReportButtonClick).RaiseCanExecuteChanged();
            }
        }
        
        public ActStatus SelectedActStatus
        {
            get { return _selectedActStatus; }
            set
            {
                _selectedActStatus = value; 
                RaisePropertyChanged(nameof(SelectedActStatus));
                ((DelegateCommand)OnViewReportButtonClick).RaiseCanExecuteChanged();
            }
        }

        public ObservableCollection<DirectionViewModel> SelectedDirections { get; }

        public Commission SelectedCommission
        {
            get { return _selectedCommission; }
            set
            {
                _selectedCommission = value; 
                RaisePropertyChanged(nameof(SelectedCommission));
                ((DelegateCommand)OnViewReportButtonClick).RaiseCanExecuteChanged();
            }
        }

        public ObservableCollection<ActStatus> ActStatuses { get; }

        public ObservableCollection<Commission> CommissionTypes { get; }

        public ObservableCollection<ChanelViewModel> Chanels { get; }



        private void LoadResultSet(IEnumerable<ActsResultSet> actsResultSets)
        {
            ((DelegateCommand)OnExportDataButtonClick).RaiseCanExecuteChanged();
            _resultTextSet = GetTextFromDataSet(actsResultSets);
        }

        private void GetDirections(IEnumerable<DirectionViewModel> selectedDirections)
        {
            SelectedDirections.Clear();
            foreach (var direction in selectedDirections)
            {
                SelectedDirections.Add(direction);
            }
        }

        #region Button Report

        private bool CanViewReport()
        {
            return (PeriodDateTime.HasValue && StartDateTime.HasValue && SelectedActStatus != null && SelectedCommission != null);
            //return true;
        }

        private void OnClickedViewReportButton()
        {
            //StringBuilder sb = new StringBuilder();
            //sb.Append('*', 10);
            //sb.Append(PeriodDateTime);
            //sb.Append('\n');
            //sb.Append(StartDateTime);
            //sb.Append('\n');
            //sb.Append(SelectedCommission.TypeDefinition);
            //sb.Append('\n');
            //sb.Append(SelectedActStatus.Name);
            //sb.Append('\n');
            //sb.Append(String.Join(" ", Chanels.Where(d => d.IsSelected).Select(d => $"{d.Code}")));
            //sb.Append('\n');
            //sb.Append('*', 10);
            //foreach (var directionViewModel in SelectedDirections)
            //{
            //    sb.Append('\n');
            //    sb.Append(directionViewModel.FullName);
            //    sb.Append('\n');
            //}
            //sb.Append('*', 10);
            //MessageBox.Show(sb.ToString());

            var userInput = new InputParametersModel
            {
                period = PeriodDateTime,
                startDate = StartDateTime,
                endDate = DateTime.Now,
                commissionType = SelectedCommission.Guid,
                statusGID = SelectedActStatus.Guid,
                SelectedChanels = Chanels.Where(d => d.IsSelected).Select(d => d.Guid),
                SelectedDirections = SelectedDirections.Select(d => d.Guid)
            };

            _eventAggregator.GetEvent<OnParametersButtonClickEvent>().Publish(userInput);

        }

        #endregion

        #region Button Export

        private bool CanClickExport()
        {
            return !String.IsNullOrEmpty(_resultTextSet);
        }

        private void OnClickedExportDataButton()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            if (saveFileDialog.ShowDialog() == true)
            {
                File.WriteAllText(saveFileDialog.FileName, _resultTextSet);
            }
        }

        #endregion



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
                Chanels.Add(new ChanelViewModel
                {
                    Code = chanel.Code,
                    Guid = chanel.Guid,
                    Name = chanel.Name.Substring(5),
                    IsSelected = false
                });
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


        private string GetTextFromDataSet(IEnumerable<ActsResultSet> resultSet)
        {
            var resultString = new StringBuilder();

            foreach (ActsResultSet actsResultSet in resultSet)
            {
                resultString.Append(actsResultSet.ToString());
                resultString.Append("\n");
            }
            return resultString.ToString();
        }

        #region INotifyDataErrorInfo implementation

        private Dictionary<string, List<string>> _errorsByPropertyName = new Dictionary<string, List<string>>();

        public bool HasErrors => _errorsByPropertyName.Any();

        public event EventHandler<DataErrorsChangedEventArgs> ErrorsChanged;

        public IEnumerable GetErrors(string propertyName)
        {
            return _errorsByPropertyName.ContainsKey(propertyName) ? _errorsByPropertyName[propertyName] : null;
        }

        private void OnErrorsChanged(string propertyName)
        {
            ErrorsChanged?.Invoke(this, new DataErrorsChangedEventArgs(propertyName));
        }

        private void AddError(string propertyName, string error)
        {
            if (!_errorsByPropertyName.ContainsKey(propertyName))
            {
                _errorsByPropertyName[propertyName] = new List<string>();
            }

            if (_errorsByPropertyName[propertyName].Contains(error)) return;
            _errorsByPropertyName[propertyName].Add(error);
            OnErrorsChanged(propertyName);
        }

        private void ClearErrors(string propertyName)
        {
            if (!_errorsByPropertyName.ContainsKey(propertyName)) return;
            _errorsByPropertyName.Remove(propertyName);
            OnErrorsChanged(propertyName);
        }


        #endregion

        #region Validation

        //private void ValidateProperty(string propertyName)
        //{
        //    ClearErrors(propertyName);
        //    ValidateProperty(nameof(PeriodDateTime));
        //    switch (propertyName)
        //    {

        //    }
        //}

        #endregion

        protected override void OnPropertyChanged(PropertyChangedEventArgs args)
        {
            base.OnPropertyChanged(args);
            _resultTextSet = null;
            ((DelegateCommand)OnExportDataButtonClick).RaiseCanExecuteChanged();
            _eventAggregator.GetEvent<ParametersChangesEvent>().Publish(true);
        }
    }
}