using Cassini.Model;
using Prism.Mvvm;

namespace Cassini.UI.ViewModel
{
    public class ActResultSetSumView: BindableBase, ISelectableViewModel
    {
        public string DirCode { get; set; }

        public int ActId { get; set; }

        public string AgentName { get; set; }

        public string INN { get; set; }

        public string DogType { get; set; }

        public decimal? SummCommission { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                RaisePropertyChanged(nameof(IsSelected));
            }
        }
    }
}