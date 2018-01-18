using System.Collections.Generic;
using Cassini.Model;
using Prism.Events;

namespace Cassini.UI.Event
{
    public class ActResultSetLoadEvent: PubSubEvent<IEnumerable<ActsResultSet>>
    {
        
    }
}