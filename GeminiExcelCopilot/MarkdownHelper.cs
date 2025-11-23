using System;
using Markdig;

namespace GeminiExcelCopilot
{
    public static class MarkdownHelper
    {
        public static string ConvertToHtml(string markdownText)
        {
            if (string.IsNullOrEmpty(markdownText)) return "";

            // 1. Markdig Pipeline'ı oluştur (Tablolar vb. için gelişmiş özellikler)
            var pipeline = new MarkdownPipelineBuilder()
                .UseAdvancedExtensions()
                .Build();

            // 2. Markdown'ı ham HTML parçasına çevir
            string rawHtml = Markdown.ToHtml(markdownText, pipeline);

            // 3. VSTO WebBrowser içinde güzel görünmesi için CSS ekle
            // (Excel'in modern görünümüne uygun stil)
            string fullHtmlDocument = $@"
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset='UTF-8'>
                    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
                    <style>
                        body {{
                            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                            font-size: 13px;
                            color: #333;
                            padding: 12px;
                            line-height: 1.6;
                            background-color: #ffffff;
                        }}
                        /* Başlıklar */
                        h1, h2, h3 {{ color: #2c3e50; margin-top: 15px; margin-bottom: 8px; }}
                        h1 {{ font-size: 18px; border-bottom: 1px solid #eee; padding-bottom: 5px; }}
                        h2 {{ font-size: 16px; }}
                        h3 {{ font-size: 14px; font-weight: bold; }}

                        /* Kod blokları */
                        pre {{
                            background-color: #f6f8fa;
                            padding: 10px;
                            border-radius: 6px;
                            overflow-x: auto;
                            border: 1px solid #e1e4e8;
                            font-family: Consolas, 'Courier New', monospace;
                        }}
                        code {{
                            background-color: #f0f2f5;
                            padding: 2px 4px;
                            border-radius: 3px;
                            font-family: Consolas, monospace;
                            color: #d63384; 
                            font-size: 12px;
                        }}
                        pre code {{
                            background-color: transparent;
                            color: #24292e;
                            padding: 0;
                        }}

                        /* Tablolar */
                        table {{
                            border-collapse: collapse;
                            width: 100%;
                            margin-bottom: 15px;
                            font-size: 12px;
                        }}
                        th, td {{
                            border: 1px solid #dfe2e5;
                            padding: 6px 10px;
                        }}
                        th {{
                            background-color: #f6f8fa;
                            font-weight: 600;
                            text-align: left;
                        }}
                        tr:nth-child(even) {{ background-color: #fcfcfc; }}
                        
                        /* Listeler */
                        ul, ol {{ padding-left: 20px; margin-bottom: 10px; }}
                        li {{ margin-bottom: 4px; }}
                    </style>
                </head>
                <body>
                    {rawHtml}
                </body>
                </html>";

            return fullHtmlDocument;
        }
    }
}