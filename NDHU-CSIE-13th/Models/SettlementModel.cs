using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NDHU_CSIE_13th.Models
{
    public class SettlementModel
    {
        public List<AdvanceModel> Advances { get; private set; }
        public List<CostModel> Costs { get; private set; }
        public List<MemberModel> Members { get; private set; }

        public SettlementModel(
            IEnumerable<AdvanceModel> advances,
            IEnumerable<CostModel> costings,
            IEnumerable<MemberModel> members)
        {
            Advances = advances.ToList();
            Costs = costings.ToList();
            Members = members.ToList();

            UpdateMemberDetial();
        }

        public void UpdateCostMemberCount()
        {
            foreach (var cost in Costs)
            {
                cost.MemberCount = Members.Count(member => member.Costs.Any(x => x.Id == cost.Id));
            }
        }

        public void UpdateMemberAmount()
        {
            foreach (var member in Members)
            {
                member.TotalCost = member.Costs.Sum(x => x.PerPersonAmount);
                member.TotalAdvance = member.Advances.Sum(x => x.Amount);
            }
        }

        private void UpdateMemberDetial()
        {
            foreach (var member in Members)
            {
                member.Costs = Costs
                    .Where(x => member.Activities.Contains(x.Desc))
                    .ToList();
                member.Advances = Advances
                    .Where(x => x.MemberName == member.Name)
                    .ToList();
            }
        }

    }
}
