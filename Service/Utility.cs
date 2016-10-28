using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AstekSuivi.Service
{
    static class Utility
    {
        public static string GetHtmlFromUrl(string url)
        {
            var html = String.Empty;

            if (String.IsNullOrEmpty(url)) return html;

            var request = (HttpWebRequest)WebRequest.Create(url);

            HttpWebResponse response = null;
            try
            {
                // get the response, to later read the stream
                response = (HttpWebResponse)request.GetResponse();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // get the response stream.
            //Stream responseStream = response.GetResponseStream();

            if (response != null)
            {
                // use a stream reader that understands UTF8
                var reader = new StreamReader(response.GetResponseStream(), encoding: Encoding.UTF8);
                html = reader.ReadToEnd();
                // close the reader
                reader.Close();
                response.Close();
            }

            return html; //return html content
        }
    }
}
