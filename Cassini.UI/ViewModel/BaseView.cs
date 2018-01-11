using System.ComponentModel;
using System.Runtime.CompilerServices;
using Cassini.UI.Annotations;

namespace Cassini.UI.ViewModel
{
    public class BaseView: INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}