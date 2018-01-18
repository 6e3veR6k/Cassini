using Prism.Mvvm;

namespace Cassini.UI.ViewModel
{
    public class ChanelViewModel:BindableBase, ISelectableViewModel
    {
        private bool _isSelected;


        public System.Guid Guid { get; set; }
        public string Name { get; set; }
        public string Code { get; set; }
        

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