using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NDHU_CSIE_13th.Models
{
    public class MemberModel
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public List<CostModel> Costs { get; set; }
        public List<AdvanceModel> Advances { get; set; }
        public List<string> Activities { get; set; }
        public int TotalCost { get; set; }
        public int TotalAdvance { get; set; }

        public MemberModel()
        {
            Costs = new List<CostModel>();
            Advances = new List<AdvanceModel>();
            Activities = new List<string>();
        }

        public string GetCostDesc()
        {
            var desc = new StringBuilder();
            var count = 1;

            foreach (var cost in Costs)
            {
                desc.AppendLine($"{count}. {cost.Item} {cost.Day}");

                count++;
            }

            return desc.ToString();
        }

        public string GetAdvanceDesc()
        {
            var desc = new StringBuilder();
            var count = 1;

            foreach (var advance in Advances)
            {
                desc.AppendLine($"{count}. {advance.Item}");

                count++;
            }

            return desc.ToString();
        }

        public int GetActualPayout()
        {
            return TotalCost > TotalAdvance ? TotalCost - TotalAdvance : 0;
        }

        public int GetActualReceive()
        {
            return TotalCost > TotalAdvance ? 0 : TotalAdvance - TotalCost;
        }
    }
}
