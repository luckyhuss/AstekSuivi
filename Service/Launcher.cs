using AstekSuivi.Model;
using System;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AstekSuivi.Service
{
    static class Launcher
    {
        private static void GetMCBFOREX(string urlForexMCB)
        {
            var content = Utility.GetHtmlFromUrl(urlForexMCB);

            var pageContent = content;
            const string patternDiv = "<div.*</div>";

            pageContent = pageContent.Replace("\n", String.Empty);
            pageContent = pageContent.Replace("\r", String.Empty);
            pageContent = pageContent.Replace("\t", String.Empty);

            // get <DIV> tags from the content, one by one
            pageContent = pageContent.Replace("</div>", "</div>" + Environment.NewLine);

            StringBuilder sbForex = new StringBuilder();
            foreach (Match m in Regex.Matches(pageContent, patternDiv, RegexOptions.IgnoreCase))
            {
                var value = m.Value;

                if (value.Contains("class=\"currency\""))
                {
                    // clear <div> and </div>
                    value = value.Replace("<div class=\"row\">", String.Empty)
                        .Replace("<div class=\"row odd\">", String.Empty)
                        .Replace("</div>", String.Empty);

                    value = value.Replace(" ", String.Empty).Replace("<spanclass=\"currency\">", String.Empty)
                        .Replace("<spanclass=\"sell\">", "Buy : ").Replace("<spanclass=\"buy\">", "Sell : ")
                        .Replace("</span>", " * ").Trim();

                    sbForex.AppendLine(value.Substring(0, value.Length - 1));
                }
            }

            MessageBox.Show(sbForex.ToString(),
                String.Format("FOREX - MCB @ {0}", DateTime.Today.ToString("dd/MM/yyyy")),
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void LaunchControl(CustomContextMenu ccm)
        {
            if (null == ccm)
            {
                return;
            }

            switch (ccm.MenuType)
            {
                case "web":
                case "app":
                case "dir":
                    try
                    {
                        Process.Start(ccm.MenuLink);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;
                case "mcb":
                    GetMCBFOREX(ccm.MenuLink);
                    break;
                default:
                    break;
            }
        }
    }
}
