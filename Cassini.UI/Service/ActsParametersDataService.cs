using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using Cassini.DA;
using Cassini.Model;

namespace Cassini.UI.Service
{
    public class ActsParametersDataService : IActsParametersDataService
    {
        private Func<CallistoDb> _contextCreator;

        public ActsParametersDataService(Func<CallistoDb> contextCreator)
        {
            _contextCreator = contextCreator;
        }

        public async Task<IEnumerable<ActStatus>> GetActStatusAsync()
        {
            using (var callistoContext = _contextCreator())
            {
                return await callistoContext.AgentActStatuses
                    .Select(s => new ActStatus
                    {
                        Guid = s.gid,
                        Name = s.Name
                    })
                    .AsNoTracking()
                    .ToListAsync();
            }
        }

        public async Task<IEnumerable<Chanel>> GetAgentChanelsAsync()
        {
            using (var callistoContext = _contextCreator())
            {
                return await callistoContext.AgentChanels
                    .Where(c => c.ParentGID != null)
                    .OrderBy(c => c.Code)
                    .Select(c => new Chanel
                    {
                        Code = c.Code,
                        Guid = c.gid,
                        Name = c.Name
                    })
                    .AsNoTracking()
                    .ToListAsync();
            }
        }

        public async Task<IEnumerable<Commission>> GetCommissionTypesAsync()
        {
            using (var callistoContext = _contextCreator())
            {
                return await callistoContext.CommissionTypes
                    .Select(t =>
                        new Commission
                        {
                            Guid = t.gid,
                            TypeDefinition = t.Caption
                        })
                    .AsNoTracking()
                    .ToListAsync();
            }
        }
    }
}