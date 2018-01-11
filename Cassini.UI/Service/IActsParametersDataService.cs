using System.Collections.Generic;
using System.Threading.Tasks;
using Cassini.Model;

namespace Cassini.UI.Service
{
    public interface IActsParametersDataService
    {
        Task<IEnumerable<ActStatus>> GetActStatusAsync();
        Task<IEnumerable<Chanel>> GetAgentChanelsAsync();
        Task<IEnumerable<Commission>> GetCommissionTypesAsync();
    }
}