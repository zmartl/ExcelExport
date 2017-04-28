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
        private readonly Color _red = Color.FromArgb(230, 43, 39);

        // GET: Statistics
        public ActionResult Index()
        {            
            return View(db.Statistics.ToList());
        }

        /// <summary>
        ///     Set the default Layout for one worksheet (Header and Footer)
        /// </summary>
        /// <param name="worksheet">actual worksheet</param>
        /// <param name="isPlanning"></param>
        private void SetTemplate(ExcelWorksheet worksheet, bool isPlanning)
        {
            try
            {
                var originalBitmap = new Bitmap(Server.MapPath("~/stadtuster.png"));
                var image = new Bitmap(originalBitmap, 135, 50);
                worksheet.HeaderFooter.OddHeader.InsertPicture(image, PictureAlignment.Left);
            }
            catch (ArgumentException e)
            {
                //TODO
            }
            if (isPlanning)
            {
                worksheet.HeaderFooter.OddHeader.RightAlignedText = "Fahrzeugplanung - Export";
                worksheet.Column(1).Width = 23;
                worksheet.Column(2).Width = 23;
                worksheet.Column(3).Width = 23;
                worksheet.Column(4).Width = 20;
            }
            else
            {
                worksheet.HeaderFooter.OddHeader.RightAlignedText = "Fahrzeugstatistik - Export";
                worksheet.Column(1).Width = 27;
                worksheet.Column(2).Width = 27;
                worksheet.Column(3).Width = 35;

            }
            worksheet.HeaderFooter.OddFooter.RightAlignedText = string.Format("{0} von {1}",
                ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);

            worksheet.PrinterSettings.Orientation = eOrientation.Portrait;

            // Change the sheet view to show it in page layout mode
            worksheet.View.PageLayoutView = true;
        }

        public int GetCountPlannedPlanningsByState(IEnumerable<Planning> plannings, string state)
        {
            var count =  plannings.Count(x => x.State.Name.Equals(state));
            return count;
        }

        public IEnumerable<Planning> GetPlannedPlannings(DateTime starTime, DateTime endTime)
        {
            var plannedCars = db.Plannings.Where(
                        x =>
                            x.StartTime <= starTime && starTime <= x.EndTime ||
                            x.StartTime <= endTime && endTime <= x.EndTime ||
                            starTime <= x.StartTime && endTime >= x.EndTime).ToList();

            return plannedCars;
        }

        /// <summary>
        ///     Create the Excel-File with all devices
        /// </summary>
        public FileInfo CreateSingleCarSheet()
        {
            bool isPlanning = false;
            Car car = new Car {Description = "BMW", Radio = "9801", Id = 1};
            DateTime starttime = DateTime.Now;
            DateTime endtime = DateTime.Now.AddDays(2);

            var filename = "Auswertung_" + car.Description + car.Radio + "_" + DateTime.Now.ToString("d") + ".xlsx";

            var response = new FileInfo(filename);

            using (var package = new ExcelPackage(response))
            {
                var worksheet = package.Workbook.Worksheets.Add("Fahrzeugstatistik");

                SetTemplate(worksheet, isPlanning);

                var allEntities = GetPlannedPlannings(starttime, endtime).Where(x => x.Car.Id == car.Id);

                var countPlannings = allEntities.Count();

                //Add the header informations
                //First row
                // COL / ROW

                worksheet.Cells[1, 1].Style.Font.Size = 18;
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
                worksheet.Cells[8, 2].Value = countPlannings;

                var col = 9;
                foreach (var state in db.States)
                {
                    worksheet.Cells[col, 1].Value = state.Name + ":";
                    worksheet.Cells[col, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Cells[col, 2].Value = GetCountPlannedPlanningsByState(allEntities, state.Name);
                    col++;
                }

                col = col + 2;
                var tableHeader = col;

                //Add the content-headers
                worksheet.Cells[tableHeader, 1].Value = "Startzeit";
                worksheet.Cells[tableHeader, 2].Value = "Endzeit";
                worksheet.Cells[tableHeader, 3].Value = "Status";

                foreach (var planning in allEntities)
                {
                    col++;
                    worksheet.Cells[col, 1].Value = planning.StartTime.ToString("d") + " " + planning.StartTime.ToString("t"); 
                    worksheet.Cells[col, 2].Value = planning.EndTime.ToString("d") + " " + planning.EndTime.ToString("t");
                    worksheet.Cells[col, 3].Value = planning.State.Name;                    
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
                    range.Style.Fill.BackgroundColor.SetColor(_red);
                    range.Style.Font.Color.SetColor(Color.White);
                    range.AutoFilter = true;
                }


                //Set property values
                package.Workbook.Properties.Subject = "Fahrzeugstatistik";
                package.Workbook.Properties.Title = "Fahrzeugstatistik";

                //Set extended property values
                package.Workbook.Properties.Company = "Verwaltungspolizei Stadtpolizei Uster";

                this.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                this.Response.AddHeader(
                          "content-disposition",
                          string.Format("attachment;  filename={0}", filename));
                this.Response.BinaryWrite(package.GetAsByteArray());
            }

            return response;
        }

        /// <summary>
        ///     Create the Excel-File with all devices
        /// </summary>
        public FileInfo CreateOverviewSheet()
        {
            bool isPlanning = true;

            DateTime starttime = DateTime.Now;
            DateTime endtime = DateTime.Now.AddDays(2);

            var filename = "Planung_" + DateTime.Now.ToString("d") + ".xlsx";

            var response = new FileInfo(filename);

            using (var package = new ExcelPackage(response))
            {
                var worksheet = package.Workbook.Worksheets.Add("Fahrzeugplanung");

                SetTemplate(worksheet, isPlanning);

                var allEntities = GetPlannedPlannings(starttime, endtime).OrderBy(x => x.StartTime);

                //Add the header informations
                //First row
                // COL / ROW
                worksheet.Cells[1, 1].Style.Font.Size = 18;
                worksheet.Cells[1, 1].Style.Font.Bold = true;
                worksheet.Cells[1, 1].Value = "Fahrzeugplanung";
                worksheet.Cells[2, 1].Style.Font.Size = 14;
                worksheet.Cells[2, 1].Value = starttime.ToString("d") + " - " + endtime.ToString("d");
                worksheet.Cells[5, 1].Value = "Erstelldatum:";
                worksheet.Cells[5, 2].Value = DateTime.Now.ToString("dd.MM.yyyy H:mm");
                worksheet.Cells[6, 1].Value = "Ersteller:";
                worksheet.Cells[6, 2].Value = "Marti, Luca";

                var col = 7;

                col = col + 2;
                var tableHeader = col;

                //Add the content-headers
                worksheet.Cells[tableHeader, 1].Value = "Startzeit";
                worksheet.Cells[tableHeader, 2].Value = "Endzeit";
                worksheet.Cells[tableHeader, 3].Value = "Fahrzeug";
                worksheet.Cells[tableHeader, 4].Value = "Status";

                foreach (var planning in allEntities)
                {
                    col++;
                    worksheet.Cells[col, 1].Value = planning.StartTime.ToString("d") + " " + planning.StartTime.ToString("t");
                    worksheet.Cells[col, 2].Value = planning.EndTime.ToString("d") + " " + planning.EndTime.ToString("t");
                    worksheet.Cells[col, 3].Value = planning.Car.Description + " - " + planning.Car.Radio;
                    worksheet.Cells[col, 4].Value = planning.State.Name;
                }

                //Format the Header
                using (var range = worksheet.Cells[4, 1, 4, 4])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                }
                var borderEnd = tableHeader - 2;
                using (var range = worksheet.Cells[borderEnd, 1, borderEnd, 4])
                {
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                }

                // Format the List-Header 
                using (var range = worksheet.Cells[tableHeader, 1, tableHeader, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(_red);
                    range.Style.Font.Color.SetColor(Color.White);
                    range.AutoFilter = true;
                }


                //Set property values
                package.Workbook.Properties.Subject = "Fahrzeugplanung";
                package.Workbook.Properties.Title = "Fahrzeugplanung";

                //Set extended property values
                package.Workbook.Properties.Company = "Verwaltungspolizei Stadtpolizei Uster";

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