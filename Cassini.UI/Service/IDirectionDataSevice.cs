using System.Collections.Generic;
using System.Threading.Tasks;
using Cassini.Model;

namespace Cassini.UI.Service
{
    public interface IDirectionDataSevice
    {
        Task<List<Direction>> GetAllAsync();
    }
}