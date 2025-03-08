using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;


using System.ComponentModel;

using System.Web.SessionState;



namespace UpLoadFile
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            foreach (string f in Request.Files.AllKeys)
            {
                HttpPostedFile file = Request.Files[f];
                file.SaveAs(@"../File/" + file.FileName);
            }
            if (Request.Params["testKey"] != null)
            {
                Response.Write(Request.Params["testKey"]);
            }

        }
    }
}

