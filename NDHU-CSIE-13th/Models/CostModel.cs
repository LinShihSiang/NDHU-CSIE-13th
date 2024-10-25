using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NDHU_CSIE_13th.Models
{
    public class CostModel
    {
        public int Id { get; set; }
        public string Item { get; set; } = string.Empty;
        public string Day { get; set; } = string.Empty;
        public int Amount {  get; set; }
        public string Desc { get; set; } = string.Empty;
        public int MemberCount { get; set; }
        public int PerPersonAmount { get { return MemberCount == 0 ? 0 : Amount / MemberCount; } }
    }
}
