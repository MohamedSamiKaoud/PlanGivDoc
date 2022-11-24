
using Microsoft.AspNetCore.Mvc;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        [HttpPost("CreateDocument")]
        public IActionResult CreateDocument(FinancialReport report)
        {
            Document doc = new Document();
            Section s = doc.AddSection();
            Paragraph par1 = new Paragraph(doc);
            Paragraph par2 = new Paragraph(doc);
            s.Paragraphs.Add(par1);
            s.Paragraphs.Add(par2);
            s.Paragraphs[0].AppendBookmarkStart("ReportTable");
            s.Paragraphs[1].AppendBookmarkEnd("ReportTable");

            //Move to bookmark
            BookmarksNavigator bn = new BookmarksNavigator(doc);
            bn.MoveToBookmark("ReportTable", true, true);
            Section section0 = doc.AddSection();

            var T1 = CreateFinancialTable(doc, section0, report);
            bn.InsertTable(T1);

            doc.Sections.Remove(section0);

            Paragraph par3 = new Paragraph(doc);
            Paragraph par4 = new Paragraph(doc);

            s.Paragraphs.Add(par3);
            s.Paragraphs.Add(par4);


            s.Paragraphs[2].AppendBookmarkStart("ContactTable");
            s.Paragraphs[3].AppendBookmarkEnd("ContactTable");

            //BookmarksNavigator bn2 = new BookmarksNavigator(doc);
            bn.MoveToBookmark("ContactTable", true, true);
            Section section1 = doc.AddSection();

            var T2 = CreateContactsTable(doc, report.DARContact.ToList(), section1);
            bn.InsertTable(T2);

            doc.Sections.Remove(section1);


            var randomName = Guid.NewGuid().ToString();
            doc.SaveToFile($"{randomName}.docx", FileFormat.Docx2013);

            return Ok(randomName);
        }

        private void CreateBookMark(Paragraph paragraph, Document doc, string bookMarkName, string bookMarkValue)
        {
            paragraph.AppendBookmarkStart(bookMarkName);
            paragraph.AppendBookmarkEnd(bookMarkName);
            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(doc);
            bookmarkNavigator.MoveToBookmark(bookMarkName);
            bookmarkNavigator.InsertText(bookMarkValue, true);

        }

        private Table CreateFinancialTable(Document doc, Section s, FinancialReport report)
        {
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
            table.TableFormat.Positioning.HorizPositionAbs = HorizontalPosition.Outside;
            table.TableFormat.Positioning.VertRelationTo = VerticalRelation.Margin;
            table.TableFormat.Positioning.VertPosition = 0;


            int index = 0;
            foreach (KeyValuePair<string, string> kvp in Header)
            {
                TableRow dataRow = table.Rows[index];
                Paragraph par = dataRow.Cells[0].AddParagraph();
                dataRow.Cells[0].CellFormat.BackColor = Color.LightSeaGreen;
                Paragraph par2 = dataRow.Cells[1].AddParagraph();
                var objValue = report.GetType().GetProperty(kvp.Key).GetValue(report, null);
                if (objValue is null)
                    objValue = string.Empty;
                CreateBookMark(par2, doc, kvp.Key, objValue.ToString());
                par.AppendText(kvp.Value);
                par.Format.HorizontalAlignment = HorizontalAlignment.Left;
                index++;
            }
            return table;
        }


        private Table CreateContactsTable(Document doc, List<DARContact> contacts, Section s)
        {
            //String[] Header = { "Informal Salutation", "Street Address", "Address2", "Full Address", "Country" };
            Dictionary<string, string> Header = new Dictionary<string, string>();
            Header.Add("InformalSalutation", "Informal Salutation");
            Header.Add("StreetAddress", "Street Address");
            Header.Add("Address2", "Address2");
            Header.Add("FullAddress", "Full Address");
            Header.Add("Country", "Country");

            Table table = s.AddTable(true);
            //table.TableFormat.Positioning.HorizPositionAbs = HorizontalPosition.Outside;
            //table.TableFormat.Positioning.VertRelationTo = VerticalRelation.Page;
            //table.TableFormat.Positioning.VertPosition = 700;
            table.ResetCells(contacts.Count + 1, Header.Count);
            TableRow FRow = table.Rows[0];
            FRow.IsHeader = true;
            FRow.Height = 23;
            FRow.RowFormat.BackColor = Color.LightSeaGreen;

            int counter = 0;
            foreach (KeyValuePair<string, string> kvp in Header)
            {
                Paragraph p = FRow.Cells[counter].AddParagraph();
                FRow.Cells[counter].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                TextRange TR = p.AppendText(kvp.Key);
                TR.CharacterFormat.FontName = "Calibri";
                TR.CharacterFormat.FontSize = 12;
                TR.CharacterFormat.Bold = true;
                counter++;
            }
            counter = 0;

            int index = 0;

            foreach (var contact in contacts)
            {
                TableRow dataRow = table.Rows[index + 1];
                foreach (KeyValuePair<string, string> kvp in Header)
                {
                    var objValue = contact.GetType().GetProperty(kvp.Key).GetValue(contact, null);
                    if (objValue is null)
                        objValue = string.Empty;
                    Paragraph p = dataRow.Cells[counter].AddParagraph();
                    CreateBookMark(p, doc, kvp.Key, objValue.ToString());
                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    counter++;
                }
                index++;
            }
            return table;

        }




    }
}