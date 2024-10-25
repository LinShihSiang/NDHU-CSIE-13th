using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using NDHU_CSIE_13th.Extensions;
using NDHU_CSIE_13th.Models;
using NDHU_CSIE_13th.Repos;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using System.Threading.Tasks;

namespace NDHU_CSIE_13th.Services
{
    public class SettlementService
    {
        private readonly ILogger<SettlementService> _logger;
        private readonly ExcelRepo _excelRepo;

        public SettlementService(
            ILogger<SettlementService> logger,
            ExcelRepo excelRepo)
        {
            _logger = logger;
            _excelRepo = excelRepo;
        }

        public void Start()
        {
            try
            {
                var settlement = _excelRepo.GetSettlement();

                settlement.UpdateCostMemberCount();
                settlement.UpdateMemberAmount();

                _excelRepo.SaveSettlement(settlement);

                var option = new JsonSerializerOptions
                {
                    Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
                    WriteIndented = true
                };

                _logger.LogInformation(JsonSerializer.Serialize(settlement, option));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Settlement Error");
            }
        }
    }
}
