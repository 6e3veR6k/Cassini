using System.Collections.Generic;
using Cassini.UI.ViewModel;
using Prism.Events;

namespace Cassini.UI.Event
{
    public class SelectedDirectionsEvent:PubSubEvent<IEnumerable<DirectionViewModel>>
    {
        
    }
}