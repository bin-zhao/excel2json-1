using System;
using System.IO;
using System.Data;
using System.Text;
using Excel;

namespace excel2json
{
    /// <summary>
    /// 应用程序
    /// </summary>
    sealed partial class Program
    {
        /// <summary>
        /// 应用程序入口
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            System.DateTime startTime = System.DateTime.Now;

            //-- 分析命令行参数
            var options = new Options();
            var parser = new CommandLine.Parser(with => with.HelpWriter = Console.Error);

            if (parser.ParseArgumentsStrict(args, options, () => Environment.Exit(-1)))
            {
                //-- 执行导出操作
                try
                {
                    Run(options);
                    //-- 程序计时
                    System.DateTime endTime = System.DateTime.Now;
                    System.TimeSpan dur = endTime - startTime;
                    Console.WriteLine(
                        string.Format("[{0}]：\t转换完成[{1}毫秒].",
                        Path.GetFileName(options.ExcelPath),
                        dur.Milliseconds)
                        );
                }
                catch (Exception exp)
                {
                    Console.WriteLine("Error: " + exp.Message);
//                     Console.WriteLine("\npress any key to continue...");
//                     Console.ReadLine();
                }
            }
        }

        /// <summary>
        /// 根据命令行参数，执行Excel数据导出工作
        /// </summary>
        /// <param name="options">命令行参数</param>
        private static void Run(Options options)
        {
            string excelPath = options.ExcelPath;
            int header = options.HeaderRows;

            // TODO 支持转换路径下的所有表格 excelDir
            if (excelPath == null || excelPath.Length <= 0)
            {
                Console.WriteLine("请输入表格文件名");
                return;
            }

            // 加载Excel文件
            using (FileStream excelFile = File.Open(excelPath, FileMode.Open, FileAccess.Read))
            {
                // Reading from a OpenXml Excel file (2007 format; *.xlsx)
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);

                // The result of each spreadsheet will be created in the result.Tables
                excelReader.IsFirstRowAsColumnNames = true;
                DataSet book = excelReader.AsDataSet();

                // 数据检测
                if (book.Tables.Count < 1)
                {
                    throw new Exception("Excel文件中没有找到Sheet: " + excelPath);
                }

                // 取得数据
                DataTable sheet = book.Tables[0];
                if (sheet.Rows.Count <= 0)
                {
                    throw new Exception("Excel Sheet中没有数据: " + excelPath);
                }

                //-- 确定编码
                Encoding cd = new UTF8Encoding(false);
                if (options.Encoding != "utf8-nobom")
                {
                    foreach (EncodingInfo ei in Encoding.GetEncodings())
                    {
                        Encoding e = ei.GetEncoding();
                        if (e.EncodingName == options.Encoding)
                        {
                            cd = e;
                            break;
                        }
                    }
                }

                //-- 导出JSON文件
                JsonExporter jsonExporter = new JsonExporter(sheet, header, options.Lowcase);
                if (options.JsonPath != null && options.JsonPath.Length > 0)
                {
                    jsonExporter.SaveToFile(options.JsonPath, cd);
                }
                else
                {
                    string jsonDir = Path.GetDirectoryName(excelPath);
                    string jsonFileName = Path.GetFileNameWithoutExtension(excelPath);
                    string jsonFilePath = Path.Combine(jsonDir, jsonFileName + ".json");
                    jsonExporter.SaveToFile(jsonFilePath, cd);
                }

                //-- 导出SQL文件
                if (options.SQLPath != null && options.SQLPath.Length > 0)
                {
                    SQLExporter exporter = new SQLExporter(sheet, header);
                    exporter.SaveToFile(options.SQLPath, cd);
                }

                //-- 生成C#定义文件
                if (options.CSharpPath != null && options.CSharpPath.Length > 0)
                {
                    string excelName = Path.GetFileName(excelPath);

                    CSDefineGenerator exporter = new CSDefineGenerator(sheet);
                    exporter.ClassComment = string.Format("// Generate From {0}", excelName);
                    exporter.SaveToFile(options.CSharpPath, cd);
                }
            }
        }
    }
}
