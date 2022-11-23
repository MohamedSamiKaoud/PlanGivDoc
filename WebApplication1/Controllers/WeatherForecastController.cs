using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading.Tasks;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        [HttpPost("CreateDocument")]
        public IActionResult CreateDocument(FinancialReport report )
        {

            Document doc = new Document();
            CreateFinancialTable(doc, report);
            var randomName = Guid.NewGuid().ToString();
            doc.SaveToFile($"{randomName}.docx", FileFormat.Docx2013);
            return Ok(randomName);
        }

        private void CreateBookMark(Document doc)
        {
            Section section = doc.AddSection();
            Paragraph paragraph = section.AddParagraph();
            paragraph.AppendText("Hello Form My First Project");
            paragraph.AppendBookmarkStart("FirstName");
            paragraph.AppendBookmarkEnd("FirstName");
            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(doc);
            bookmarkNavigator.MoveToBookmark("FirstName");
            bookmarkNavigator.InsertText("Mohamed Agour", true);
            //doc.SaveToFile(@"C:\Users\mohamed.agour\Desktop\MR.Berj\Output.docx", FileFormat.Docx2013);

        }

        private void CreateFinancialTable(Document doc, FinancialReport report)
        {
            Section s = doc.AddSection();

            Dictionary<string, string> Header = new Dictionary<string, string>();
            Header.Add("FundCode", "Fund Code");
            Header.Add("FundName", "Fund Name");
            Header.Add("StartPeriod", "Start Period");
            Header.Add("EndPeriod", "End Period");
            Header.Add("OpeningFundValue", "Opening Fund Value");
            Header.Add("FundNetContrbution", "Fund Net Contrbution");
            Header.Add("AssessmentForAdmin", "Assessment For Admin");
            Header.Add("NetInvestmentReturn", "Net Investment Return");
            Header.Add("GrantsFromFund", "Grants From Fund");
            Header.Add("TransfersToCharitableGiftFund", "Transfers To Charitable GiftFund");
            Header.Add("ClosingValue", "Closing Value");
            Header.Add("OpeningBalanceGrantMoney", "Opening Balance Grant Money");
            Header.Add("OpeningUnrestrictedCapitalBalance", "Opening Unrestricted Capital Balance");
            Header.Add("ClosingBalanceGrantMoney", "Closing Balance Grant Money");
            Header.Add("ClosingUnrestrictedCapitalBalance", "Closing Unrestricted Capital Balance");
            Header.Add("TotalGlGifts", "Total GlGifts");
            Header.Add("TotalGrants", "Total Grants");

            Table table = s.AddTable(true);
            table.ResetCells(Header.Count, 2);
            table.TableFormat.WrapTextAround = true;
            table.TableFormat.Positioning.VertPosition = 43;

            int index = 0;
            foreach (KeyValuePair<string, string> kvp in Header)
            {
                TableRow dataRow = table.Rows[index];
                Paragraph par = dataRow.Cells[0].AddParagraph();
                dataRow.Cells[0].CellFormat.BackColor = Color.LightSeaGreen;
                Paragraph par2 = dataRow.Cells[1].AddParagraph();
                var objValue = report.GetType().GetProperty(kvp.Key).GetValue(report, null);
                par.AppendText(kvp.Value);
                par2.AppendText(objValue.ToString());
                par.Format.HorizontalAlignment = HorizontalAlignment.Left;
                index++;
            }

        }



    }
}