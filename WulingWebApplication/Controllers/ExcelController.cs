
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using WulingWebApplication.DAL;

namespace WulingWebApplication.Controllers
{
    public class ExcelController : Controller
    {
        //
        // GET: /foot/
        //private static readonly String Folder = "/files";
        public ActionResult Index()
        {
            return View();
        }

        /// 导入excel文档
        public async Task<ActionResult> ImportExcel(HttpPostedFileBase[] files)
        {
            //1.接收客户端传过来的数据
            //HttpPostedFileBase file = Request.Files["file"];
            HttpPostedFileBase file = files[0];
            if (file == null || file.ContentLength <= 0)
            {
                return Json("请选择要上传的Excel文件", JsonRequestBehavior.AllowGet);
            }
            //string filepath = Server.MapPath(Folder);
            //if (!Directory.Exists(filepath))
            //{
            //  Directory.CreateDirectory(filepath);
            //}
            //var fileName = Path.Combine(filepath, Path.GetFileName(file.FileName));
            // file.SaveAs(fileName);
            //获取一个streamfile对象，该对象指向一个上传文件，准备读取改文件的内容
            Stream streamfile = file.InputStream;
            DataTable dt = new DataTable();
            string FinName = Path.GetExtension(file.FileName);
            if (FinName != ".xls" && FinName != ".xlsx")
            {
                return Json("只能上传Excel文档", JsonRequestBehavior.AllowGet);
            }
            else
            {
                try
                {
                    if (FinName == ".xls")
                    {
                        //创建一个webbook，对应一个Excel文件(用于xls文件导入类)
                        HSSFWorkbook hssfworkbook = new HSSFWorkbook(streamfile);
                        dt = await ExcelDAL.ImExport(dt, hssfworkbook);
                    }
                    else
                    {
                        XSSFWorkbook hssfworkbook = new XSSFWorkbook(streamfile);
                        dt = await ExcelDAL.ImExport(dt, hssfworkbook);
                    }
                    return Json("", JsonRequestBehavior.AllowGet);
                }
                catch (Exception ex)
                {
                    return Json("导入失败 ！" + ex.Message, JsonRequestBehavior.AllowGet);
                }
            }

        }

    }
}