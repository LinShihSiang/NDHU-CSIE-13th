using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NDHU_CSIE_13th.Models
{
    public class AdvanceModel
    {
        public string MemberName { get; set; } = string.Empty;
        public string Date { get; set; } = string.Empty;   
        public string Item { get; set; } = string.Empty;
        public int Amount { get; set; }

        public AdvanceModel()
        {
        }
    }
}
