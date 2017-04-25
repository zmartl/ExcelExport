using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Web.Mvc;

using ExcelTest.Models;

using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

namespace ExcelTest.Controllers
{
    public class StatisticsController : Controller
    {
        private readonly ExcelTestContext db = new ExcelTestContext();
        private readonly ExcelPackage p = new ExcelPackage();
        private readonly Color _blue = Color.FromArgb(230, 43, 39);

        // GET: Statistics
        public ActionResult Index()
        {            
            return View(db.Statistics.ToList());
        }        

        private Stream DownloadImage(string url)
        {
            try
            {
                WebRequest req = WebRequest.Create(url);
                WebResponse response = req.GetResponse();
                Stream stream = response.GetResponseStream();
                return stream;
            }
            catch (Exception)
            {
                //TODO
            }
            return null;
        }        

        /// <summary>
        ///     Set the default Layout for one worksheet for Airport Zurich (Header and Footer)
        /// </summary>
        /// <param name="worksheet">actual worksheet</param>
        private void SetTemplate(ExcelWorksheet worksheet)
        {
            try
            {
                var originalBitmap = new Bitmap("C:\\Development\\ExcelTest\\ExcelTest\\stadtuster.png");
                var image = new Bitmap(originalBitmap, 135, 50);
                worksheet.HeaderFooter.OddHeader.InsertPicture(image, PictureAlignment.Left);
            }
            catch (ArgumentException e)
            {
                //TODO
            }
            worksheet.HeaderFooter.OddHeader.RightAlignedText = "Fahrzeugstatistik - Export";
            worksheet.HeaderFooter.OddFooter.RightAlignedText = string.Format("{0} von {1}",
                ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
            worksheet.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FileName;

            worksheet.PrinterSettings.Orientation = eOrientation.Landscape;

            // Change the sheet view to show it in page layout mode
            worksheet.View.PageLayoutView = true;
        }

        private int GetCountPlanningsByState(string state)
        {
            var res = db.Plannings.Where(x => x.State.Name.Equals(state));
            return res.Count();
        }

        /// <summary>
        ///     Create the Excel-File with all devices
        /// </summary>
        public FileInfo CreateSheet()
        {
            Car car = new Car {Description = "BMW", Radio = "9801"};
            DateTime starttime = DateTime.Now;
            DateTime endtime = DateTime.Now.AddDays(2);

            var filename = "Auswertung_" + car.Description + car.Radio + "_" + DateTime.Now.ToString("d") + ".xlsx";

            var response = new FileInfo(filename);

            using (var package = new ExcelPackage(response))
            {
                var worksheet = package.Workbook.Worksheets.Add("Fahrzeugstatistik");

                SetTemplate(worksheet);

                var allEntities = db.Plannings.ToList();

                var countDevices = allEntities.Count;


                //Add the header informations
                //First row
                // COL / ROW
                worksheet.Cells[1, 1].Style.Font.Size = 16;
                worksheet.Cells[1, 1].Style.Font.Bold = true;
                worksheet.Cells[1, 1].Value = "Auswertung - " + car.Description + " " + car.Radio;
                worksheet.Cells[2, 1].Style.Font.Size = 14;
                worksheet.Cells[2, 1].Value = starttime.ToString("d") + " - " + endtime.ToString("d");
                worksheet.Cells[5, 1].Value = "Erstelldatum:";
                worksheet.Cells[5, 2].Value = DateTime.Now.ToString("dd.MM.yyyy H:mm");
                worksheet.Cells[6, 1].Value = "Ersteller:";
                worksheet.Cells[6, 2].Value = "Marti, Luca";

                //Second row
                worksheet.Cells[8, 1].Style.Font.Bold = true;
                worksheet.Cells[8, 1].Value = "Anzahl Einträge:";
                worksheet.Cells[8, 2].Style.Font.Bold = true;
                worksheet.Cells[8, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[8, 2].Value = countDevices;

                var col = 9;
                foreach (var state in db.States)
                {
                    worksheet.Cells[col, 1].Value = state.Name + ":";
                    worksheet.Cells[col, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Cells[col, 2].Value = GetCountPlanningsByState(state.Name);
                    col++;
                }

                col = col + 2;
                var tableHeader = col;

                //Add the content-headers
                worksheet.Cells[tableHeader, 1].Value = "Startzeit";
                worksheet.Cells[tableHeader, 2].Value = "Endzeit";
                worksheet.Cells[tableHeader, 3].Value = "Status";

                foreach (var device in allEntities)
                {
                    col++;
                    worksheet.Cells[col, 1].Value = device.StartTime.ToString("d") + " " + device.StartTime.ToString("t"); 
                    worksheet.Cells[col, 2].Value = device.EndTime.ToString("d") + " " + device.EndTime.ToString("t");
                    worksheet.Cells[col, 3].Value = device.State.Name;                    
                }

                //Format the Header
                using (var range = worksheet.Cells[4, 1, 4, 3])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                }
                var borderEnd = tableHeader-2;
                using (var range = worksheet.Cells[borderEnd, 1, borderEnd, 3])
                {
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                }

                // Format the List-Header 
                using (var range = worksheet.Cells[tableHeader, 1, tableHeader, 3])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(_blue);
                    range.Style.Font.Color.SetColor(Color.White);
                    range.AutoFitColumns();
                    range.AutoFilter = true;
                    range.AutoFitColumns(1);
                    range.AutoFitColumns(2);
                    range.AutoFitColumns(3);
                }

                //Set property values
                package.Workbook.Properties.Subject = "Gerätelebenslauftool";
                package.Workbook.Properties.Title = "Geräteinventarliste";

                //Set extended property values
                package.Workbook.Properties.Company = "Flughafen Zürich AG";

                //package.Save();

                this.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                this.Response.AddHeader(
                          "content-disposition",
                          string.Format("attachment;  filename={0}", filename));
                this.Response.BinaryWrite(package.GetAsByteArray());
            }

            return response;
        }


        // GET: Statistics/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            var statistic = db.Statistics.Find(id);
            if (statistic == null)
                return HttpNotFound();
            return View(statistic);
        }

        // GET: Statistics/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Statistics/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(
            [Bind(Include = "StatisticId,StartDate,EndDate,CreationDate,Creator")] Statistic statistic)
        {
            if (ModelState.IsValid)
            {
                db.Statistics.Add(statistic);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(statistic);
        }

        // GET: Statistics/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            var statistic = db.Statistics.Find(id);
            if (statistic == null)
                return HttpNotFound();
            return View(statistic);
        }

        // POST: Statistics/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(
            [Bind(Include = "StatisticId,StartDate,EndDate,CreationDate,Creator")] Statistic statistic)
        {
            if (ModelState.IsValid)
            {
                db.Entry(statistic).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(statistic);
        }

        // GET: Statistics/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            var statistic = db.Statistics.Find(id);
            if (statistic == null)
                return HttpNotFound();
            return View(statistic);
        }

        // POST: Statistics/Delete/5
        [HttpPost]
        [ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            var statistic = db.Statistics.Find(id);
            db.Statistics.Remove(statistic);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                db.Dispose();
            base.Dispose(disposing);
        }
    }
}