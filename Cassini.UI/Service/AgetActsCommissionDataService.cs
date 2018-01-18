using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Threading.Tasks;
using Cassini.DA;
using Cassini.UI.ViewModel;
using System.Linq;
using Cassini.Model;

namespace Cassini.UI.Service
{
    public class AgetActsCommissionDataService : IAgetActsCommissionDataService
    {
        private Func<CallistoDb> _contextCreator;

        public AgetActsCommissionDataService(Func<CallistoDb> contextCreator)
        {
            _contextCreator = contextCreator;
        }

        public async Task<IEnumerable<ActsResultSet>> GetActsComissionsSumResults(InputParametersModel userParametersModel)
        {
            using (var contextCreator = _contextCreator())
            {
                return await contextCreator.AgentActsComissionsSum(userParametersModel.period, userParametersModel.startDate,
                    DateTime.Now, userParametersModel.statusGID, userParametersModel.commissionType)
                    .Join(userParametersModel.SelectedDirections, a => a.DirectionGid, d => d, (a, d) => a)
                    .Join(userParametersModel.SelectedChanels, a => a.ChanelGID, c => c, (a, c) => a)
                    .Join(contextCreator.AgentChanels, a => a.ChanelGID, ac => ac.gid, (a, ac) => new ActsResultSet
                    {
                        IdentificationCodeEDRPOU = a.IdentificationCodeEDRPOU,
                        ActId = a.ActId,
                        AgentName = a.AgentName,
                        BranchCode = a.BranchCode,
                        ChanelName = ac.Name,
                        CommissionValue = a.CommissionValue,
                        DocumentType = a.DocumentType,
                        ProgramCode = a.ProgramCode,
                        RealPaymentValue = a.RealPaymentValue
                    })
                    .AsNoTracking().ToListAsync();
            }
        }

    }
}