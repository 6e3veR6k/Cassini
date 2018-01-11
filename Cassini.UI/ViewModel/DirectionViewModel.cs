using Cassini.Model;

namespace Cassini.UI.ViewModel
{
    public class DirectionViewModel: BaseView, ISelectableViewModel
    {
        private bool _isSelected;
        public System.Guid Guid { get; set; }
        public string Code { get; set; }
        public string Title { get; set; }

        public string FullName => $"{this.Code} {this.Title}";

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public DirectionViewModel(Direction direction)
        {
            Guid = direction.Guid;
            Code = direction.Code;
            Title = direction.Title;
            IsSelected = false;
        }
    }
}