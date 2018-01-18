using System;
using System.ComponentModel;

namespace Cassini.Model
{
    public class ActsResultSet
    {
        [DisplayName("Код ЄДРПОУ")]
        public string IdentificationCodeEDRPOU { get; set; }

        [DisplayName("Назва агента")]
        public string AgentName { get; set; }

        [DisplayName("Програма")]
        public string ProgramCode { get; set; }

        [DisplayName("Платіж")]
        public decimal? RealPaymentValue { get; set; }


        [DisplayName("Комісія")]
        
        public decimal? CommissionValue { get; set; }

        [DisplayName("Відділення")]
        public string BranchCode { get; set; }

        [DisplayName("Канал")]
        public string ChanelName { get; set; }

        [DisplayName("Тип документа")]
        public string DocumentType { get; set; }

        [DisplayName("Номер акту")]
        public int ActId { get; set; }

        public override string ToString()
        {
            return
                $"{this.IdentificationCodeEDRPOU}\t{this.AgentName}\t{ProgramCode}\t{RealPaymentValue}\t{CommissionValue}\t{BranchCode}\t{ChanelName}\t{DocumentType}\t{ActId}";
        }
    }
}