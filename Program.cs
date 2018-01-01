using Data.Models;
using Data.Service;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace Data
{
    class Program
    {
        static void Main(string[] args)
        {
            ExportToDb();
            //ExportToExecl();

        }

        private static void ExportToExecl()
        {

            string sFileName = $"{Guid.NewGuid()}.xlsx";
            //FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            FileInfo file = new FileInfo(sFileName);

            using (ApplicationDbContext db = new ApplicationDbContext())
            {
                List<String> buildNameList = (from e in db.Exports
                                                orderby e.BuildingName
                                                select e.BuildingName).Distinct().ToList();
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    int i = 0;

                    foreach (string buildName in buildNameList)
                    {
                        // 添加worksheet
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(buildName);

                        List<Export> userInfos = (from e in db.Exports
                                                    where e.BuildingName == buildName
                                                    orderby e.Unit
                                                    select e).ToList();
                        List<List<Export>> table = new List<List<Export>>();

                        List<Export> cols = new List<Export>();
                        int u = userInfos[0].Unit;
                        int f = 7;
                        int r = 2;

                        foreach (Export userInfo in userInfos)
                        {
                            if (userInfo.Unit != u)
                            {
                                table.Add(cols);
                                u = userInfo.Unit;
                                cols = new List<Export>();
                            }
                            cols.Add(userInfo);
                            if (userInfo.Floor > f) f = userInfo.Floor;
                            if (userInfo.Room % 100 > r) r = userInfo.Room % 100;
                        }
                        table.Add(cols);
                        u = table[table.Count - 1][0].Unit;

                        if (u < 3) u = 3;

                        foreach (List<Export> exports in table)
                        {
                            i += exports.Count;
                        }
                        Console.WriteLine(i);

                        int colStart = 2;
                        int rowStart = 4;

                        string[] colsName = {"房号", "用户姓名", "固话", "宽带", "ITV", "移动手机", "联通手机", "电信手机"};
                        int index = colsName.Length;

                        //添加头

                        worksheet.Cells[1, 1, 1, u * index + 1].Merge = true;
                        worksheet.Cells[1, 1, 1, u * index + 1].Value = table[0][0].BuildingArea + "小区—楼宇通信状况登记表";
                        worksheet.Cells[1, 1, 1, u * index + 1].Style.Font.Bold = true;
                        worksheet.Cells.Style.Font.Size = 12;
                        worksheet.Cells[1, 1, 1, u * index + 1].Style.Font.Size = 16;

                        worksheet.Cells.Style.Font.Name = "宋体";
                        worksheet.Row(1).Height = 20.4;
                        worksheet.Row(2).Height = 15.6;
                        worksheet.Row(3).Height = 15.6;

                        for (int j = 4; j <= rowStart - 1 + f * r; j++)
                        {
                            worksheet.Row(j).Height = 42.6;
                        }

                        for (int j = 1; j <= colStart + (u * index); j++)
                        {
                            worksheet.Column(j).Width = 10;
                        }

                        worksheet.Cells["A2:A3"].Merge = true;
                        worksheet.Cells["A2:A3"].Value = "楼号";

                        for (int c = 0; c < u; c++)
                        {
                            //第2行的头
                            worksheet.Cells[rowStart - 2, colStart + (c * index) + 0, rowStart - 2,
                                colStart + (c * index) + index - 1].Merge = true;
                            worksheet.Cells[rowStart - 2, colStart + (c * index) + 0, rowStart - 1,
                                    colStart + (c * index) + index - 1].Value = "（ " + (c + 1) + " ）单元";

                            //第3行的头
                            for (int cs = 0; cs < index; cs++)
                            {
                                worksheet.Cells[rowStart - 1, colStart + (c * index) + cs].Value = colsName[cs];
                            }
                            //第4行
                            for (int j = 0; j < f; j++)
                            {
                                for (int k = 0; k < r; k++)
                                {
                                    worksheet.Cells[rowStart + j * r + k, colStart + (c * index)].Value =
                                        (j + 1) * 100 + k + 1;
                                    if (table[0][0].BuildingNo != 0)
                                    {
                                        worksheet.Cells[rowStart + j * r + k, 1].Value = table[0][0].BuildingNo;
                                    }
                                    
                                }
                            }
                        }

                        for (int m = 2; m <= rowStart - 1 + f * r; m++)
                        {
                            
                            for (int n = 1; n < colStart + (u * index); n++)
                            {
                                worksheet.Cells[m,n].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            }
                        }

                        // u f r
                        //添加内容
                        foreach (List<Export> exports in table)
                        {
                            foreach (Export export in exports)
                            {
                                int row = rowStart - 1 + (export.Floor - 1) * r + export.Room % 100;

                                int col = colStart + (export.Unit - 1) * index;

                                Console.WriteLine("(" + worksheet.Cells[2, col].Value + "," +
                                                  worksheet.Cells[row, col].Value + ")===>" + export.Address);

                                if (worksheet.Cells[row,col].Value.Equals(export.Room))
                                {
                                    worksheet.Cells[row, col + 1].Value = export.Name;
                                    worksheet.Cells[row, col + 2].Value = export.Call;
                                    worksheet.Cells[row, col + 3].Value = export.BrandWidth;
                                    worksheet.Cells[row, col + 4].Value = export.ITV;
                                    worksheet.Cells[row, col + 5].Value = export.MobilePhone;
                                    worksheet.Cells[row, col + 6].Value = export.LinkPhone;
                                    worksheet.Cells[row, col + 7].Value = export.TelePhone;
                                }
                                else
                                {
                                    Console.WriteLine(export.Address);
                                    Console.WriteLine(worksheet.Cells[row, col].Value + "-------->" + export.Room);
                                }
                            }
                        }

                        

                        worksheet.Cells.Style.WrapText = true;//自动换行
                        worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }
                    package.Save();

                    Console.WriteLine(i);
                }
            }
        }


        static void ExportToDb()
        {
            Regex regexNum = new Regex(@"[0-9]+");

            Regex regexBuildingArea = new Regex(@".*小区|.*社区|.*家属楼|.*家属院|.*综合楼|.*住宅楼|.*商住楼|.*公寓楼");

            Regex regexBuildingName = new Regex(@".*楼");

            Regex regexBuildingNo = new Regex(@"[0-9]+号楼");

            Regex regexUnit = new Regex(@"[0-9]+单元");

            Regex regexFloor = new Regex(@"[0-9]+层");

            Regex regexRoom = new Regex(@"层[0-9]+");

            Regex regexDirectionUnit = new Regex(@"[东西左右中]单元");

            Regex regexDirectionRoom = new Regex(@"[东西左右中][户门手]");

            //string str = "武威市凉州区靶场法院东侧设计院家属楼2单元5层502";

            using (ApplicationDbContext db = new ApplicationDbContext())
            {
                //List<Info> infos = db.Infos.Where(info => EF.Functions.Like(info.装机地址, "%楼%单元%层%")).ToList();
                List<Info> infos = db.Infos.Where(info => EF.Functions.Like(info.装机地址, "%[单元号]%") 
                                                          && !EF.Functions.Like(info.装机地址, "%铺%")
                                                          && !EF.Functions.Like(info.装机地址, "%办公楼%")
                                                          && !EF.Functions.Like(info.装机地址, "%物业楼%")
                                                          && !EF.Functions.Like(info.装机地址, "%门诊楼%")
                                                          && !EF.Functions.Like(info.装机地址, "%部队%")
                                                          && !EF.Functions.Like(info.装机地址, "%餐厅%")
                                                          && !EF.Functions.Like(info.装机地址, "%批发%")).ToList();
                /*
                 * select * from Infos 
                   where [装机地址] LIKE '%[单元号]%'
                   AND [装机地址] NOT LIKE '%铺%' 
                   AND [装机地址] NOT LIKE '%办公楼%' 
                   AND [装机地址] NOT LIKE '%物业楼%' 
                   AND [装机地址] NOT LIKE '%门诊楼%'
                   AND [装机地址] NOT LIKE '%部队%'
                   AND [装机地址] NOT LIKE '%餐厅%'
                 */
                foreach (Info info in infos)
                {
                    Export export = new Export {Address = info.装机地址};


                    if (!string.IsNullOrEmpty(info.CUST_NAME) && info.CUST_NAME != "#N/A")
                    {
                        export.Name = info.CUST_NAME;
                    }

                    if (!string.IsNullOrEmpty(info.宽带账号) && info.宽带账号 != "#N/A")
                    {
                        export.BrandWidth = info.宽带账号;
                    }

                    if (!string.IsNullOrEmpty(info.关联ITV账号) && info.关联ITV账号 != "#N/A")
                    {
                        export.ITV = info.关联ITV账号;
                    }

                    //-----------------------电话-------------------------------------------

                    if (Regex.IsMatch(info.用户联系方式, @"^(?:133|153|1700|1701|1702|177|173|18[019])\d{7,8}$"))
                    {
                        export.TelePhone = info.用户联系方式;
                    }
                    else if (Regex.IsMatch(info.用户联系方式, @"^(?:13[0-2]|145|15[56]|176|1704|1707|1708|1709|171|18[56])\d{7,8}$"))
                    {
                        export.LinkPhone = info.用户联系方式;
                    }
                    else if (Regex.IsMatch(info.用户联系方式, @"^134[0-8]\d{7}$|^(?:13[5-9]|147|15[0-27-9]|178|1703|1705|1706|18[2-478])\d{7,8}$"))
                    {
                        export.MobilePhone = info.用户联系方式;
                    }
                    else
                    {
                        export.Call = info.用户联系方式;
                    }

                    //-----------------------电话 END---------------------------------------


                    //-----------------------地址-------------------------------------------
                    

                    if (regexBuildingNo.IsMatch(info.装机地址))
                    {
                        string buildingNo = regexBuildingNo.Match(info.装机地址).ToString();
                        export.BuildingName = regexBuildingName.Match(info.装机地址).ToString().Trim();
                        MatchCollection ms = regexNum.Matches(buildingNo);
                        export.BuildingNo = Convert.ToInt32(ms[ms.Count - 1].ToString());
                    }
                    else if (regexBuildingName.IsMatch(info.装机地址))
                    {
                        export.BuildingName = regexBuildingName.Match(info.装机地址).ToString().Trim();
                    }

                    if (regexBuildingArea.IsMatch(info.装机地址))
                    {
                        export.BuildingArea = regexBuildingArea.Match(info.装机地址).ToString().Trim();
                    }
                    else if(regexBuildingNo.IsMatch(info.装机地址))
                    {
                        export.BuildingArea = info.装机地址.Substring(0,
                            info.装机地址.IndexOf(export.BuildingNo + "号楼", StringComparison.Ordinal));
                    }
                    else
                    {
                        export.BuildingArea = export.BuildingName;
                    }

                    if (regexUnit.IsMatch(info.装机地址))
                    {
                        string unit = regexUnit.Match(info.装机地址).ToString();
                        MatchCollection ms = regexNum.Matches(unit);
                        export.Unit = Convert.ToInt32(ms[ms.Count - 1].ToString());
                    }
                    else if(regexDirectionUnit.IsMatch(info.装机地址))
                    {
                        string unit = regexDirectionUnit.Match(info.装机地址).ToString();
                        if (new Regex(@"[西左]").IsMatch(unit))
                        {
                            export.Unit = 1;
                        }
                        else if(new Regex(@"[右东]").IsMatch(unit))
                        {
                            export.Unit = 3;
                        }
                        else if (new Regex(@"[中]").IsMatch(unit))
                        {
                            export.Unit = 2;
                        }
                        
                    }

                    if (regexFloor.IsMatch(info.装机地址))
                    {
                        string floor = regexFloor.Match(info.装机地址).ToString();
                        MatchCollection ms = regexNum.Matches(floor);
                        export.Floor = Convert.ToInt32(ms[ms.Count - 1].ToString());
                    }

                    if (regexRoom.IsMatch(info.装机地址))
                    {
                        string room = regexRoom.Match(info.装机地址).ToString();
                        MatchCollection ms = regexNum.Matches(room);
                        export.Room = Convert.ToInt32(ms[ms.Count - 1].ToString());
                    }
                    else if (regexDirectionRoom.IsMatch(info.装机地址))
                    {
                        string room = regexDirectionRoom.Match(info.装机地址).ToString();
                        if (new Regex(@"[西左]").IsMatch(room))
                        {
                            export.Room = export.Floor * 100 + 1;
                        }
                        else if (new Regex(@"[右东]").IsMatch(room))
                        {
                            export.Room = export.Floor * 100 + 2;
                        }
                    }

                    //-----------------------地址 END-------------------------------------------


                    db.Exports.Add(export);
                }

                db.SaveChanges();
            }
        }
    }
}
