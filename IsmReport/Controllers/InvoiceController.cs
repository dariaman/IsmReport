using System;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Web.Mvc;
using IsmReport.Models;
using Microsoft.Office.Interop.Word;
using word = Microsoft.Office.Interop.Word;

namespace IsmReport.Controllers
{
    public class InvoiceController : Controller
    {
        private ApplicationDbContext _context = new ApplicationDbContext();

        // GET: Invoice
        public ActionResult Index()
        {
            if (!(TempData["pesanJudul"]==null || TempData["pesanType"] == null || TempData["pesanText"] == null))
            {
                ViewBag.pesanJudul = TempData["pesanJudul"].ToString();
                ViewBag.pesanType = TempData["pesanType"].ToString();
                ViewBag.pesanText = TempData["pesanText"].ToString();
            }
            return View(_context.InvoiceModel.ToList());
        }

        //// GET: Invoice/Details/5
        //public ActionResult Details(int? id)
        //{
        //    if (id == null)
        //    {
        //        return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
        //    }
        //    InvoiceModel invoiceModel = _context.InvoiceModel.Find(id);
        //    if (invoiceModel == null)
        //    {
        //        return HttpNotFound();
        //    }
        //    return View(invoiceModel);
        //}

        //// GET: Invoice/Create
        //public ActionResult Create()
        //{
        //    return View();
        //}

        //// POST: Invoice/Create
        //// To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        //// more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        //[HttpPost]
        //[ValidateAntiForgeryToken]
        //public ActionResult Create([Bind(Include = "id,InvoiceNo,InvoiceDate,PeriodeBln,PeriodeThn,Deskripsi,Qty,GrandTotal,Status,CreateDate,UpdateDate")] InvoiceModel invoiceModel)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        db.InvoiceModel.Add(invoiceModel);
        //        db.SaveChanges();
        //        return RedirectToAction("Index");
        //    }

        //    return View(invoiceModel);
        //}

        //// GET: Invoice/Edit/5
        //public ActionResult Edit(int? id)
        //{
        //    if (id == null)
        //    {
        //        return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
        //    }
        //    InvoiceModel invoiceModel = db.InvoiceModel.Find(id);
        //    if (invoiceModel == null)
        //    {
        //        return HttpNotFound();
        //    }
        //    return View(invoiceModel);
        //}

        //// POST: Invoice/Edit/5
        //// To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        //// more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        //[HttpPost]
        //[ValidateAntiForgeryToken]
        //public ActionResult Edit([Bind(Include = "id,InvoiceNo,InvoiceDate,PeriodeBln,PeriodeThn,Deskripsi,Qty,GrandTotal,Status,CreateDate,UpdateDate")] InvoiceModel invoiceModel)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        db.Entry(invoiceModel).State = EntityState.Modified;
        //        db.SaveChanges();
        //        return RedirectToAction("Index");
        //    }
        //    return View(invoiceModel);
        //}

        //// GET: Invoice/Delete/5
        //public ActionResult Delete(int? id)
        //{
        //    if (id == null)
        //    {
        //        return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
        //    }
        //    InvoiceModel invoiceModel = db.InvoiceModel.Find(id);
        //    if (invoiceModel == null)
        //    {
        //        return HttpNotFound();
        //    }
        //    return View(invoiceModel);
        //}

        //// POST: Invoice/Delete/5
        //[HttpPost, ActionName("Delete")]
        //[ValidateAntiForgeryToken]
        //public ActionResult DeleteConfirmed(int id)
        //{
        //    InvoiceModel invoiceModel = db.InvoiceModel.Find(id);
        //    db.InvoiceModel.Remove(invoiceModel);
        //    db.SaveChanges();
        //    return RedirectToAction("Index");
        //}

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _context.Dispose();
            }
            base.Dispose(disposing);
        }

        public ActionResult GenerateInv()
        {
            var cmd = _context.Database.Connection.CreateCommand();
            cmd.Parameters.Clear();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "Generate_Invoice";
            cmd.Parameters.Clear();
            try
            {
                cmd.Connection.Open();
                cmd.ExecuteNonQuery();

                TempData["pesanJudul"] = "Success!";
                TempData["pesanType"] = "alert-success";
                TempData["pesanText"] = "Generate invoice sukses";
            }
            catch(Exception ex)
            {
                TempData["pesanJudul"] = "Error!";
                TempData["pesanType"] = "alert-danger";
                TempData["pesanText"] = ex.Message;
            }
            finally
            {
                cmd.Connection.Close();
            }
            

            return RedirectToAction("Index");
        }

        public ActionResult Printinv(int id)
        {
            var datainv = new InvoiceModel();
            try
            {
                var cmd = _context.Database.Connection.CreateCommand();
                cmd.Parameters.Clear();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = @"select * from invoice iv where iv.id=@id";
                cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = id });
                try
                {
                    cmd.Connection.Open();
                    var result = cmd.ExecuteReader();
                    while (result.Read())
                    {
                        datainv.id = (int)result["id"];
                        datainv.InvoiceNo = result["InvoiceNo"].ToString();
                        datainv.InvoiceDate = result["InvoiceDate"].ToString() == "" ? (DateTime?)null : (DateTime?)result["InvoiceDate"];
                        datainv.PeriodeBln = result["PeriodeBln"].ToString();
                        datainv.PeriodeThn = result["PeriodeThn"].ToString();
                        datainv.Deskripsi = result["Deskripsi"].ToString();
                        datainv.Status = result["Status"].ToString();
                        datainv.Filename = result["Filename"].ToString();
                        datainv.Qty = (int)result["Qty"];
                        datainv.GrandTotal = (Decimal)result["GrandTotal"];
                        datainv.CreateDate = (DateTime)result["CreateDate"];
                        datainv.UpdateDate = result["UpdateDate"].ToString() == "" ? (DateTime?)null : (DateTime?)result["UpdateDate"];
                    }
                    if (datainv == null || datainv.id == 0) throw new Exception("Data Invoice tidak ditemukan");
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                finally
                {
                    cmd.Dispose();
                    cmd.Connection.Close();
                }

                if (!(datainv.Qty > 0) && (datainv.GrandTotal > 0)) throw new Exception("Nilai tagihan nol");

                var tmstamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds().ToString();
                var invfile = String.Concat("INV", datainv.PeriodeThn.Trim(), datainv.PeriodeBln.Trim(), tmstamp.Trim(), ".pdf");

                if (datainv.Filename == "" || datainv.InvoiceDate == null)
                {
                    datainv.Filename = invfile;
                    // update tgl Invoice
                    datainv.InvoiceDate = UpdateInvoicePrint(id, datainv.Filename);
                    CreatePdfInvoice(datainv);
                }

                string pathTemp = ConfigurationManager.AppSettings["FileInvoice"];
                Response.Headers.Add("content-disposition", "attachment; filename=" + datainv.Filename);
                return File(new FileStream(pathTemp + datainv.Filename, FileMode.Open), "application/pdf");

            }
            catch (Exception ex)
            {
                ViewBag.pesanJudul = "Error!";
                ViewBag.pesanType = "alert-danger";
                ViewBag.pesanText = ex.Message;
                //throw new Exception(ex.Message);
                return RedirectToAction("Index");
            }
        }

        private void CheckInvoicePrint(int idInvoice)
        {
            var datainv = new InvoiceModel();
            try
            {
                var cmd = _context.Database.Connection.CreateCommand();
                cmd.Parameters.Clear();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = @"select * from invoice iv where iv.id=@id";
                cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = idInvoice });
                try
                {
                    cmd.Connection.Open();
                    var result = cmd.ExecuteReader();
                    while (result.Read())
                    {
                        datainv.id = (int)result["id"];
                        datainv.InvoiceDate = result["InvoiceDate"].ToString() == "" ? (DateTime?)null : (DateTime?)result["InvoiceDate"];
                        datainv.Filename = result["Filename"].ToString();
                    }
                    if (datainv == null || datainv.id == 0)
                    {
                        throw new Exception("Data Invoice tidak ditemukan");
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                finally
                {
                    cmd.Dispose();
                    cmd.Connection.Close();
                }

            }
            catch (Exception ex)
            {
                ViewBag.pesanJudul = "Error!";
                ViewBag.pesanType = "alert-danger";
                ViewBag.pesanText = ex.Message;
                //throw new Exception(ex.Message);
            }
        }

        private void CreatePdfInvoice(InvoiceModel datainv)
        {
            string pathTemp = ConfigurationManager.AppSettings["FileInvoice"];
            string pathDoc = ConfigurationManager.AppSettings["TemplateInvoice"];

            var matchCase = true;
            var matchWholeWord = true;
            var matchWildcards = false;
            var matchSoundsLike = false;
            var matchAllWordForms = false;
            var forward = true;
            var wrap = 1;
            var format = false;
            var replace = 2;

            // Parameter di MS Word
            var InvoiceNo = "{NOINVOICE}";
            var TglInvoice = "{TGLinvoice}";
            var deskripsi = "{deskripsi}";
            var qty = "{qty}";
            var Total = "{total}";
            var grandTotal = "{grandtotal}";

            Application app = new word.Application();
            app.Visible = false;
            Document doc = app.Documents.Open(pathDoc);
            try
            {
                app.Selection.Find.Execute(InvoiceNo, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, datainv.InvoiceNo, replace);
                app.Selection.Find.Execute(TglInvoice, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, (datainv.InvoiceDate ?? DateTime.Now).ToString("dd MMM yyyy"), replace);
                app.Selection.Find.Execute(deskripsi, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, datainv.Deskripsi, replace);
                app.Selection.Find.Execute(qty, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, datainv.Qty.ToString("#,##0"), replace);
                app.Selection.Find.Execute(Total, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, datainv.GrandTotal.ToString("#,##0"), replace);
                app.Selection.Find.Execute(grandTotal, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, datainv.GrandTotal.ToString("#,##0"), replace);

                doc.SaveAs2(pathTemp + datainv.Filename, word.WdSaveFormat.wdFormatPDF);
            }
            catch (Exception ex)
            {
                //ViewBag.pesanJudul = "Error!";
                //ViewBag.pesanType = "alert-danger";
                //ViewBag.pesanText = ex.Message;
                throw new Exception(ex.Message);
            }
            finally
            {
                doc.Close(word.WdSaveOptions.wdDoNotSaveChanges);
                app.Quit();
            }
        }

        private DateTime UpdateInvoicePrint(int idInvoice, string FileName)
        {
            var TglPrint = DateTime.Now;
            var cmd = _context.Database.Connection.CreateCommand();
            cmd.Parameters.Clear();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = @"update invoice set InvoiceDate=@tgl,Filename=@file,UpdateDate=@tgl where id=@id";
            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = idInvoice });
            cmd.Parameters.Add(new SqlParameter("@tgl", SqlDbType.DateTime) { Value = TglPrint });
            cmd.Parameters.Add(new SqlParameter("@file", SqlDbType.VarChar) { Value = FileName });

            try
            {
                cmd.Connection.Open();
                cmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                cmd.Connection.Close();
            }
            return TglPrint;
        }

    }
}
