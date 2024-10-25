using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using NDHU_CSIE_13th.Enums;
using NDHU_CSIE_13th.Extensions;
using NDHU_CSIE_13th.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NDHU_CSIE_13th.Repos
{
    public class ExcelRepo : IDisposable
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<ExcelRepo> _logger;
        private readonly List<SheetSettingModel> _sheetSettings;
        private XLWorkbook _workbook;

        public ExcelRepo(
            IConfiguration configuration,
            ILogger<ExcelRepo> logger)
        {
            _configuration = configuration;
            _logger = logger;
            _sheetSettings = _configuration.GetSection("Sheet").Get<List<SheetSettingModel>>();
        }

        ~ExcelRepo() 
        {
            Dispose();
        }

        public SettlementModel GetSettlement()
        {
            LoadExcel();

            return new SettlementModel(
                GetAdvances(),
                GetCosts(),
                GetMembers());
        }

        public void SaveSettlement(SettlementModel settlement)
        {
            var filePath = _configuration["OutputPath"].ToString();
            var fileName = $"{Path.GetFileNameWithoutExtension(_configuration["FileName"].ToString())}_結算.xlsx";
            var fileFullName = Path.Combine(filePath, fileName);

            if(!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }

            using (var workbook = new XLWorkbook())
            {
                AddAdvanceSheet(workbook, settlement.Advances);
                AddCostSheet(workbook, settlement.Costs);
                AddMembersSheet(workbook, settlement.Members);

                workbook.SaveAs(fileFullName);
            }
        }

        public void Dispose()
        {
            _workbook.Dispose();
        }


        private List<AdvanceModel> GetAdvances() 
        {
            var worksheet = GetWorkSheet(SheetTypeEnum.Substitute);
            var result = new List<AdvanceModel>();

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                result.Add(new AdvanceModel()
                {
                    MemberName = row.Cell(1).GetValue<string>(),
                    Date = row.Cell(2).GetValue<string>(),
                    Item = row.Cell(3).GetValue<string>(),
                    Amount = Convert.ToInt32(row.Cell(4).GetValue<string>()),
                });
            }

            return result;
        }

        private List<CostModel> GetCosts()
        {
            var worksheet = GetWorkSheet(SheetTypeEnum.Spending);
            var result = new List<CostModel>();

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                result.Add(new CostModel()
                {
                    Id = Convert.ToInt32(row.Cell(1).GetValue<string>()),
                    Item = row.Cell(2).GetValue<string>(),
                    Day = row.Cell(3).GetValue<string>(),
                    Amount = Convert.ToInt32(row.Cell(4).GetValue<string>()),
                    Desc = row.Cell(5).GetValue<string>(),
                });
            }

            return result;
        }

        private List<MemberModel> GetMembers()
        {
            var worksheet = GetWorkSheet(SheetTypeEnum.Member);
            var result = new List<MemberModel>();

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                result.Add(new MemberModel()
                {
                    Id = Convert.ToInt32(row.Cell(1).GetValue<string>()),
                    Name = row.Cell(2).GetValue<string>(),
                    Activities = row.Cell(3).GetValue<string>()
                    .Split(',')
                    .Select(x => x.TrimStart())
                    .ToList(),
                });
            }

            return result;
        }

        private IXLWorksheet GetWorkSheet(SheetTypeEnum type) 
        {
            var sheet = _sheetSettings.First(x => x.Type == type);

            return _workbook.Worksheets.First(x => x.Name == sheet.Name);
        }

        private void LoadExcel()
        {
            if (_workbook != null) return;

            var filePath = _configuration["FilePath"].ToString();
            var fileName = _configuration["FileName"].ToString();
            var fileFullName = Path.Combine(filePath, fileName);

            if (!File.Exists(fileFullName))
            {
                throw new FileNotFoundException();
            }

            _workbook = new XLWorkbook(fileFullName);
        }



        private void AddAdvanceSheet(XLWorkbook workbook, List<AdvanceModel> advances)
        {
            var worksheet = workbook.AddWorksheet("代墊項目");

            worksheet.SetHeaders(new List<string>()
            {
                "代墊人員",
                "日期",
                "項目",
                "金額"
            });

            var rowNum = 2;

            foreach (var advance in advances)
            {
                worksheet.Cell(rowNum, 1).Value = advance.MemberName;
                worksheet.Cell(rowNum, 2).Value = advance.Date;
                worksheet.Cell(rowNum, 3).Value = advance.Item;
                worksheet.Cell(rowNum, 4).Value = advance.Amount;

                rowNum++;
            }

            worksheet.Columns().Width = 20;
        }

        private void AddCostSheet(XLWorkbook workbook, List<CostModel> costs)
        {
            var worksheet = workbook.AddWorksheet("開銷明細");

            worksheet.SetHeaders(new List<string>()
            {
                "Id",
                "項目",
                "第 N 天",
                "金額",
                "參與人數",
                "每人平均"
            });

            var rowNum = 2;

            foreach (var cost in costs)
            {
                worksheet.Cell(rowNum, 1).Value = cost.Id;
                worksheet.Cell(rowNum, 2).Value = cost.Item;
                worksheet.Cell(rowNum, 3).Value = cost.Day;
                worksheet.Cell(rowNum, 4).Value = cost.Amount;
                worksheet.Cell(rowNum, 5).Value = cost.MemberCount;
                worksheet.Cell(rowNum, 6).Value = cost.PerPersonAmount;

                rowNum++;
            }

            worksheet.Columns().Width = 20;
            worksheet.Column(1).Width = 5;
        }

        private void AddMembersSheet(XLWorkbook workbook, List<MemberModel> members)
        {
            foreach (var member in members)
            {
                var worksheet = workbook.AddWorksheet($"{member.Name}-收支明細");

                worksheet.SetHeaders(new List<string>()
                {
                    "姓名",
                    "參與項目",
                    "代墊項目",
                    "支出總金額",
                    "代墊總金額",
                    "應給付",
                    "應收款"
                });

                var rowNum = 2;

                worksheet.Cell(rowNum, 1).Value = member.Name;
                worksheet.Cell(rowNum, 2).Value = member.GetCostDesc();
                worksheet.Cell(rowNum, 3).Value = member.GetAdvanceDesc();
                worksheet.Cell(rowNum, 4).Value = member.TotalCost;
                worksheet.Cell(rowNum, 5).Value = member.TotalAdvance;
                worksheet.Cell(rowNum, 6).Value = member.GetActualPayout();
                worksheet.Cell(rowNum, 7).Value = member.GetActualReceive();

                worksheet.Cells().Style.Alignment.WrapText = true;
                worksheet.Cells().Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                worksheet.Cells().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                worksheet.Columns().Width = 20;
                worksheet.Column(2).Width = 40;
                worksheet.Column(3).Width = 40;
            }
        }

    }
}
