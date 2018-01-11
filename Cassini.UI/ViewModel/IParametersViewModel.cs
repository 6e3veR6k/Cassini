using System.Threading.Tasks;

namespace Cassini.UI.ViewModel
{
    public interface IParametersViewModel
    {
        Task GetActStatusAsync();
        Task GetAgentChanelsAsync();
        Task GetCommissionTypesAsync();
    }
}