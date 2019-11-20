using System.IO;
using System.Web;
using System.Web.Hosting;
using System.Web.Mvc;

namespace Testing
{
    public static class CssLinkHelper
    {        
        public static IHtmlString StyleSheet(this HtmlHelper helper, string stylesheetname)
        {
            if (helper is null)
            {
                throw new System.ArgumentNullException(nameof(helper));
            }
            // define the virtual path to the css file (see note below)
            var virtualpath = "~/" + stylesheetname;
            // get the real path to the css file
            var realpath = HostingEnvironment.MapPath(virtualpath);
            // get the file info of the css file
            var fileinfo = new FileInfo(realpath);

            // create a full (virtual) path to the css file including a cache busting parameter (e.g. /main.css?12345678)
            var outputpath = VirtualPathUtility.ToAbsolute(virtualpath) + "?" + fileinfo.LastWriteTime.ToFileTime();
            // define the link tag for the style sheet
            var tagdefinition = string.Format("<link rel=\"stylesheet\" type=\"text/css\" href=\"{0}\" />", outputpath);

            // return html string of the tag
            return new HtmlString(tagdefinition);
        }
    }
}