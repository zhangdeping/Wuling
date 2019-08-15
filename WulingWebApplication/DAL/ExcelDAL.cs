using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using WulingWebApplication.Models;

namespace WulingWebApplication.DAL
{
    public class ExcelDAL
    {

   
        ///<summary>
        /// #region 两种不同版本的操作excel
        /// 扩展名*.xlsx
        /// </summary>
        public static async Task<DataTable> ImExport(DataTable dt, IWorkbook hssfworkbook)
        {
            NPOI.SS.UserModel.ISheet sheet = hssfworkbook.GetSheetAt(0);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            for (int j = 0; j < (sheet.GetRow(0).LastCellNum); j++)
            {
                dt.Columns.Add(sheet.GetRow(0).Cells[j].ToString());
            }
            while (rows.MoveNext())
            {
                XSSFRow row = (XSSFRow)rows.Current;
                DataRow dr = dt.NewRow();
                for (int i = 0; i < row.LastCellNum; i++)
                {
                    NPOI.SS.UserModel.ICell cell = row.GetCell(i);
                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        dr[i] = cell.ToString();
                    }
                }
                dt.Rows.Add(dr);
            }
            dt.Rows.RemoveAt(0);

            #region 往数据库表添加数据
            using (WuLinEntities1 db = new WuLinEntities1())
            {
                if (dt != null && dt.Rows.Count != 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string 时间 = dt.Rows[i]["时间"].ToString();
                        string 国产进口 = dt.Rows[i]["国产/进口"].ToString();
                        string 省 = dt.Rows[i]["省"].ToString();
                        string 市 = dt.Rows[i]["市"].ToString();
                        string 县 = dt.Rows[i]["县"].ToString();
                        string 制造商 = dt.Rows[i]["制造商"].ToString();
                        string 车辆型号 = dt.Rows[i]["车辆型号"].ToString();
                        string 品牌 = dt.Rows[i]["品牌"].ToString();
                        string 车型 = dt.Rows[i]["车型"].ToString();
                        string 排量 = dt.Rows[i]["排量"].ToString();
                        string 变速器 = dt.Rows[i]["变速器"].ToString();
                        string 车辆类型 = dt.Rows[i]["车辆类型"].ToString();
                        string 车身型式 = dt.Rows[i]["车身型式"].ToString();
                        string 燃油类型 = dt.Rows[i]["燃油类型"].ToString();
                        string 使用性质 = dt.Rows[i]["使用性质"].ToString();
                        string 所有权 = dt.Rows[i]["所有权"].ToString();
                        string 抵押标记 = dt.Rows[i]["抵押标记"].ToString();
                        string 性别 = dt.Rows[i]["性别"].ToString();
                        string 年龄 = dt.Rows[i]["年龄"].ToString();
                        string 车身颜色 = dt.Rows[i]["车身颜色"].ToString();
                        string 发动机型号 = dt.Rows[i]["发动机型号"].ToString();
                        string 功率 = dt.Rows[i]["功率"].ToString();
                        string 排放标准 = dt.Rows[i]["排放标准"].ToString();
                        string 轴距 = dt.Rows[i]["轴距"].ToString();
                        string 轮胎规格 = dt.Rows[i]["轮胎规格"].ToString();
                        string 车外廓长 = dt.Rows[i]["车外廓长"].ToString();
                        string 车外廓宽 = dt.Rows[i]["车外廓宽"].ToString();
                        string 车外廓高 = dt.Rows[i]["车外廓高"].ToString();
                        string 准确排量 = dt.Rows[i]["准确排量"].ToString();
                        string 核定载客 = dt.Rows[i]["核定载客"].ToString();
                        string 总质量 = dt.Rows[i]["总质量"].ToString();
                        string 整备质量 = dt.Rows[i]["整备质量"].ToString();
                        string 轴数 = dt.Rows[i]["轴数"].ToString();
                        string 前轮距 = dt.Rows[i]["前轮距"].ToString();
                        string 后轮距 = dt.Rows[i]["后轮距"].ToString();
                        string 保有量 = dt.Rows[i]["保有量"].ToString(); ;
                        //int.TryParse(dt.Rows[i]["保有量"].ToString(), out 保有量);

                        PassengerVehicle pv = new PassengerVehicle();
                        pv.Id = Guid.NewGuid();
                        pv.使用性质 = 使用性质;
                        pv.保有量 = Convert.ToInt32(保有量);
                        pv.准确排量 = (准确排量);
                        pv.制造商 = 制造商;
                        pv.前轮距 = (前轮距);
                        pv.功率 = (功率);
                        pv.县 = 县;
                        pv.发动机型号 = 发动机型号;
                        pv.变速器 = 变速器;
                        pv.后轮距 = (后轮距);
                        pv.品牌 = 品牌;
                        pv.国产进口 = 国产进口;
                        pv.市 = 市;
                        pv.年龄 = (年龄);
                        pv.性别 = 性别;
                        pv.总质量 = (总质量);
                        pv.所有权 = 所有权;
                        pv.抵押标记 = 抵押标记;
                        pv.排放标准 = 排放标准;
                        pv.排量 = (排量);
                        pv.整备质量 = (整备质量);
                        pv.时间 = 时间;
                        pv.核定载客 = (核定载客);
                        pv.燃油类型 = 燃油类型;
                        pv.省 = 省;
                        pv.车型 = 车型;
                        pv.车外廓宽 = (车外廓宽);
                        pv.车外廓长 = (车外廓长);
                        pv.车外廓高 = (车外廓高);
                        pv.车身型式 = 车身型式;
                        pv.车身颜色 = 车身颜色;
                        pv.车辆型号 = 车辆型号;
                        pv.车辆类型 = 车辆类型;
                        pv.轮胎规格 = 轮胎规格;
                        pv.轴数 = (轴数);
                        pv.轴距 = (轴距);

                        db.PassengerVehicles.Add(pv);
                        try
                        {
                            await db.SaveChangesAsync();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("第"+i+"条："+e.Message);
                        }


                    }
                   


                }
                #endregion
            }

            return dt;
        }

       // #region 两种不同版本的操作excel
       // ///<summary>
       // /// 扩展名*.xls
       // /// </summary>
       // public static DataTable ImExport(DataTable dt, HSSFWorkbook hssfworkbook)
       // {
       //     // 在webbook中添加一个sheet,对应Excel文件中的sheet,取出第一个工作表，索引是0 
       //     NPOI.SS.UserModel.ISheet sheet = hssfworkbook.GetSheetAt(0);
       //     System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
       //     for (int j = 0; j < (sheet.GetRow(0).LastCellNum); j++)
       //     {
       //         dt.Columns.Add(sheet.GetRow(0).Cells[j].ToString());
       //     }
       //     while (rows.MoveNext())
       //     {
       //         HSSFRow row = (HSSFRow)rows.Current;
       //         DataRow dr = dt.NewRow();
       //         for (int i = 0; i < row.LastCellNum; i++)
       //         {
       //             NPOI.SS.UserModel.ICell cell = row.GetCell(i);
       //             if (cell == null)
       //             {
       //                 dr[i] = null;
       //             }
       //             else
       //             {
       //                 dr[i] = cell.ToString();
       //             }
       //         }
       //         dt.Rows.Add(dr);
       //     }
       //     dt.Rows.RemoveAt(0);
       //     if (dt != null && dt.Rows.Count != 0)
       //     {
       //         for (int i = 0; i < dt.Rows.Count; i++)
       //         {
       //             string categary = dt.Rows[i]["页面"].ToString();
       //             string fcategary = dt.Rows[i]["分类"].ToString();
       //             string fTitle = dt.Rows[i]["标题"].ToString();
       //             string fUrl = dt.Rows[i]["链接"].ToString();
       //             WuLinDAL.AddRecorder(categary, fcategary, fTitle, fUrl);
       //         }

       //     }
       //     return dt;
       // }
       //#endregion
    }
}