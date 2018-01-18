using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using Cassini.DA;
using Cassini.Model;

namespace Cassini.UI.Service
{
    public class DirectionDataSevice : IDirectionDataSevice
    {
        private Func<CallistoDb> _contextCreator;

        public DirectionDataSevice(Func<CallistoDb> contextCreator)
        {
            _contextCreator = contextCreator;
        }

        public async Task<List<Direction>> GetAllAsync()
        {
            using (var callistoContext = _contextCreator())
            {
                return await callistoContext.GetDirections()
                    .OrderBy(b => b.BranchCode)
                    .Select(b =>
                        new Direction {Code = b.BranchCode, Title = b.Name, Guid = b.gid}).ToListAsync();
            }
        }
    }
}
