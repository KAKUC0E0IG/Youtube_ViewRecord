using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.Apis.YouTube.v3;
using Google.Apis.Services;
using OfficeOpenXml;

namespace ViewCount
{
    class Program
    {
        //輸入Google Cloud 的API金鑰
        private const string ApiKey = "你的金鑰";
        //輸入想取得觀看數的影片ID
        private static readonly string[] VideoId = new string[] { "", "", "", "" };
        static void Main(string[] args)
        {
            //輸入Excel檔案儲存路徑
            string excelFilePath = @"C:\result.xlsx";
            foreach (string video in VideoId)
            {
                var (viewCount, title) = GetYouTubeViewCount(video).Result;
                SaveToExcel(title, viewCount, excelFilePath);
            }
        }
        //取得YouTube影片觀看數
        private static async Task<(long ViewCount, string Title)> GetYouTubeViewCount(string videoId)
        {
            var youtubeService = new YouTubeService(new BaseClientService.Initializer()
            {
                ApiKey = ApiKey,
                ApplicationName = "YouTuberViewer"
            });
            //取得影片的觀看數和標題
            var requset = youtubeService.Videos.List("statistics,snippet");
            requset.Id = videoId;
            //執行請求
            var response = await requset.ExecuteAsync();
            var video = response.Items.FirstOrDefault();
            //如果影片存在且有觀看數
            if (video != null && video.Statistics != null)
            {
                long viewCount = (long)(video.Statistics.ViewCount ?? 0);
                string title = video.Snippet.Title;
                return (viewCount, title);
            }
            return (0, string.Empty);
        }
        //儲存到Excel
        private static void SaveToExcel(string title, long viewCount, string excelFilePath)
        {
            //如果檔案存在就讀取，不存在就建立新的
            FileInfo file = new FileInfo(excelFilePath);
            ExcelPackage package;
            if (file.Exists)
            {
                using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    package = new ExcelPackage(stream);
                }
            }
            else
            {
                package = new ExcelPackage();
            }
            //如果工作表不存在就建立新的
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == title);
            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add(title);
                worksheet.Cells[1, 1].Value = "觀看次數";
                worksheet.Cells[1, 2].Value = "記錄時間";
                worksheet.Cells[1, 3].Value = "一天成長量";
                worksheet.Cells[1, 4].Value = "一週成長量";
            }
            //取得目前的列數
            int row = worksheet.Dimension?.Rows + 1 ?? 2;
            worksheet.Cells[row, 1].Value = viewCount;
            worksheet.Cells[row, 2].Value = DateTime.Now.ToString("yyyy/MM/dd");
            //計算一天成長量和一週成長量
            if (row > 3)
            {
                worksheet.Cells[row, 3].Value = (viewCount - Convert.ToInt32(worksheet.Cells[row - 1, 1].Value)).ToString();
                //每七天計算一次
                if ((row - 1) % 7 == 0)
                {
                    int weektotal = 0;
                    for (int i = 0; i < 7; i++)
                    {
                        weektotal += Convert.ToInt32(worksheet.Cells[row - i, 1].Value);
                    }
                    worksheet.Cells[row, 4].Value = weektotal.ToString();
                }
            }
            //儲存檔案
            using (var stream = new FileStream(excelFilePath, FileMode.Create))
            {
                package.SaveAs(stream);
            }
        }
    }
}
