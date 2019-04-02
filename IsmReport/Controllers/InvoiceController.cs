using System;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Web.Mvc;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using IsmReport.Models;
using Microsoft.Office.Interop.Word;
using word = Microsoft.Office.Interop.Word;

namespace IsmReport.Controllers
{
    public class InvoiceController : Controller
    {
        private ApplicationDbContext _context = new ApplicationDbContext();

        public ActionResult Index()
        {
            if (!(TempData["pesanJudul"] == null || TempData["pesanType"] == null || TempData["pesanText"] == null))
            {
                ViewBag.pesanJudul = TempData["pesanJudul"].ToString();
                ViewBag.pesanType = TempData["pesanType"].ToString();
                ViewBag.pesanText = TempData["pesanText"].ToString();
            }
            return View(_context.InvoiceModel.ToList());
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
            catch (Exception ex)
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
                    TempData["pesanJudul"] = "Error!";
                    TempData["pesanType"] = "alert-danger";
                    TempData["pesanText"] = ex.Message;
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
                    try
                    {
                        CreatePdfInvoice(datainv);
                        // update tgl Invoice
                        datainv.InvoiceDate = UpdateInvoicePrint(id, datainv.Filename);
                    }
                    catch (Exception ex)
                    {
                        TempData["pesanJudul"] = "Error!";
                        TempData["pesanType"] = "alert-danger";
                        TempData["pesanText"] = ex.Message;
                        return RedirectToAction("Index");
                    }
                }

                string pathTemp = ConfigurationManager.AppSettings["FileInvoice"];

                string Fileinvpdf = pathTemp + datainv.Filename;
                if (!System.IO.File.Exists(Fileinvpdf))
                {
                    TempData["pesanJudul"] = "Error!";
                    TempData["pesanType"] = "alert-danger";
                    TempData["pesanText"] = "File template invoice tidak ditemukan " + Fileinvpdf;
                    return RedirectToAction("Index");
                }

                Response.Headers.Add("content-disposition", "attachment; filename=" + datainv.Filename);
                return File(new FileStream(Fileinvpdf, FileMode.Open), "application/octet-stream");

            }
            catch (Exception ex)
            {
                TempData["pesanJudul"] = "Error!";
                TempData["pesanType"] = "alert-danger";
                TempData["pesanText"] = ex.Message;
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

        //private void CreatePdfInvoice2(InvoiceModel datainv)
        //{
        //    using (WordprocessingDocument doc = WordprocessingDocument.Open(@"D:\DocTemplate\INVOICE ISOMEDIK (002).docx", true))
        //    {
        //        var body = doc.MainDocumentPart.Document.Body;
        //        var paras = body.Elements<Paragraph>();

        //        foreach (var para in paras)
        //        {
        //            foreach (var run in para.Elements<Run>())
        //            {
        //                foreach (var text in run.Elements<Text>())
        //                {
        //                    if (text.Text.Contains("text-to-replace")) text.Text = text.Text.Replace("text-to-replace", "replaced-text");
        //                }
        //            }
        //        }
        //    }

        //}

        private void CreatePdfInvoice(InvoiceModel datainv)
        {
            string pathTemp = ConfigurationManager.AppSettings["FileInvoice"];
            string pathDoc = ConfigurationManager.AppSettings["TemplateInvoice"];

            if (!Directory.Exists(pathTemp)) throw new Exception("Directori " + pathTemp + " tidak ditemukan ");
            if (!System.IO.File.Exists(pathDoc)) throw new Exception("File template invoice tidak ditemukan " + pathDoc);

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
