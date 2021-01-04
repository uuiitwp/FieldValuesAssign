using LinqToExcel;
using LinqToExcel.Query;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

namespace FieldValuesAssign
{
    // Token: 0x02000008 RID: 8
    public class ExcelHelper
    {
        // Token: 0x06000035 RID: 53 RVA: 0x00003C38 File Offset: 0x00001E38
        public ExcelHelper()
        {
            string dllPath = Assembly.GetExecutingAssembly().Location;
            string _directory = Path.GetDirectoryName(dllPath);
            string _xlsxPath = Path.Combine(_directory, "template.xlsx");
            string roaming = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            directory = Path.Combine(roaming, "FieldValuesAssign");
            DirectoryInfo dirinfo = new DirectoryInfo(directory);
            if (!dirinfo.Exists)
            {
                dirinfo.Create();
            }
            xlsxPath = Path.Combine(directory, "template.xlsx");
            FileInfo fileinfo = new FileInfo(xlsxPath);
            if (!fileinfo.Exists)
            {
                File.Copy(_xlsxPath, xlsxPath);
            }
            loadConfig();
            loadData();
        }

        // Token: 0x06000036 RID: 54 RVA: 0x00003C84 File Offset: 0x00001E84
        public void open()
        {
            Process.Start(xlsxPath);
        }

        // Token: 0x06000037 RID: 55 RVA: 0x00003CB8 File Offset: 0x00001EB8
        public void loadConfig()
        {
            ExcelQueryFactory f = new ExcelQueryFactory(xlsxPath);
            ExcelQueryable<Row> q = f.Worksheet("系统配置");
            Dictionary<string, string> r = (from row in q
                                            select row).ToList<Row>().ToDictionary((Row x) => x["键"].ToString(), (Row y) => y["值"].ToString());
            config = r;
            f.Dispose();
        }

        // Token: 0x06000038 RID: 56 RVA: 0x00003D60 File Offset: 0x00001F60
        public void loadData()
        {
            ExcelQueryFactory f = new ExcelQueryFactory(xlsxPath);
            string sheetName = config["表名"];
            ExcelQueryable<Row> q = f.Worksheet(sheetName);
            List<Row> r = q.ToList<Row>();
            data = r;
            f.Dispose();
        }

        // Token: 0x06000039 RID: 57 RVA: 0x00003DD0 File Offset: 0x00001FD0
        public int getListBoxCount()
        {
            string searchKey = "字段个数";
            IEnumerable<int> r = from row in config
                                 where row.Key == searchKey
                                 select int.Parse(row.Value);
            if (r.Count<int>() == 0)
            {
                throw new ArgumentNullException(string.Format("在系统配置表中找不到\"{0}\"字段!", searchKey));
            }
            return r.First();
        }

        // Token: 0x0600003A RID: 58 RVA: 0x00003E4C File Offset: 0x0000204C
        public int getFieldCount()
        {
            return getListBoxCount();
        }

        // Token: 0x0600003B RID: 59 RVA: 0x00004120 File Offset: 0x00002320
        public Dictionary<int, string> getFieldNames()
        {
            string searchKey = "字段";
            var r = from row in config
                    where row.Key.StartsWith(searchKey) && int.TryParse(row.Key.Replace(searchKey, ""), out int temp)
                    select new
                    {
                        Value = row.Value,
                        index = int.Parse(row.Key.Replace(searchKey, ""))
                    };
            if (r.Count() == 0)
            {
                throw new ArgumentNullException(string.Format("在系统配置表中找不到\"{0}\"字段!", searchKey));
            }
            var r2 = from row in r
                     orderby row.index
                     select new
                     {
                         row.index,
                         row.Value
                     };
            return r2.ToDictionary(e => e.index, e => e.Value);
        }

        // Token: 0x0600003C RID: 60 RVA: 0x0000427C File Offset: 0x0000247C
        public string getFieldsType()
        {
            string searchKey = "表达式类型";
            IEnumerable<string> r = (from row in config
                                     where row.Key.Trim() == searchKey
                                     select row).Select(delegate (KeyValuePair<string, string> row)
                                     {
                                         if ("PYTHON" == row.Value.ToUpper())
                                         {
                                             return "PYTHON";
                                         }
                                         if (!("VB" == row.Value.ToUpper()))
                                         {
                                             return "STRING";
                                         }
                                         return "VB";
                                     });
            if (r.Count<string>() == 0)
            {
                throw new ArgumentNullException(string.Format("在系统配置表中找不到\"{0}\"字段!", searchKey));
            }
            return r.First<string>();
        }

        // Token: 0x0600003D RID: 61 RVA: 0x000043C4 File Offset: 0x000025C4
        public Dictionary<int, string[]> getFieldtoField()
        {
            string searchKey = "赋值字段";
            var r = from row in config
                    where row.Key.StartsWith(searchKey) && int.TryParse(row.Key.Replace(searchKey, ""), out int temp)
                    select new
                    {
                        Value = row.Value.Split(new char[]
                {
                    ' ',
                    ',',
                    '，',
                    ';',
                    '；'
                }, StringSplitOptions.RemoveEmptyEntries),
                        index = int.Parse(row.Key.Replace(searchKey, ""))
                    };
            if (r.Count() == 0)
            {
                throw new ArgumentNullException(string.Format("在系统配置表中找不到\"{0}\"字段!", searchKey));
            }
            var r2 = from row in r
                     orderby row.index
                     select new
                     {
                         row.index,
                         row.Value
                     };
            return r2.ToDictionary(e => e.index, e => e.Value);
        }

        // Token: 0x04000020 RID: 32
        private readonly string directory;

        // Token: 0x04000021 RID: 33
        private readonly string xlsxPath;

        // Token: 0x04000022 RID: 34
        private Dictionary<string, string> config;

        // Token: 0x04000023 RID: 35
        public List<Row> data;
    }
}
