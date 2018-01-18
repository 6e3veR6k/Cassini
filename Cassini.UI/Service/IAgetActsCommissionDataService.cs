using System.Collections.Generic;
using System.Threading.Tasks;
using Cassini.DA;
using Cassini.Model;
using Cassini.UI.ViewModel;

namespace Cassini.UI.Service
{
    public interface IAgetActsCommissionDataService
    {
        Task<IEnumerable<ActsResultSet>> GetActsComissionsSumResults(InputParametersModel userParametersModel);
    }
}