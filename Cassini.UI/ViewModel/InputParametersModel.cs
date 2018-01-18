using System;
using System.Collections.Generic;

namespace Cassini.UI.ViewModel
{
    public class InputParametersModel
    {
        public DateTime? period;
        public DateTime? startDate;
        public DateTime? endDate;
        public Guid? statusGID;
        public Guid? commissionType;
        public IEnumerable<Guid> SelectedDirections;
        public IEnumerable<Guid> SelectedChanels;
    }
}