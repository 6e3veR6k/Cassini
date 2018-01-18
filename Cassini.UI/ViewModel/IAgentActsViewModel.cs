using System.Threading.Tasks;

namespace Cassini.UI.ViewModel
{
    public interface IAgentActsViewModel
    {
        Task GetCommissionActs(InputParametersModel inputParameters);
    }
}