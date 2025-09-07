using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;

namespace XizheC
{
    public static class PathConfig
    {
     public static string TemplatePath 
    {
        get 
        { //使用包装类来调的原因而不是直接使用是因为若多个地方都是调了它，若不使用包装类配置文件里名称修改，那么多个cs文件都要修改名称，使用包装类就
            //只有包装类修改一次即可。
            return ConfigurationManager.AppSettings["Balance Sheet"].ToString(); 
        }
    }
    
  

    }
}
