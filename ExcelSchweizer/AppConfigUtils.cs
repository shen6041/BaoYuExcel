using System;
using System.Configuration;

namespace MatchStall
{
    // ===============================================================================
    // Class Name          :    AppConfigUtils
    // Class Version       :    V02C01
    // Author              :    shenjunguo
    // Create Time         :    2013/4/9 14:54:44
    // Update Time         :    2013/4/9 14:54:44
    // Description         :    
    // ===============================================================================
    // Copyright © 天闻数媒. All rights reserved.
    // ===============================================================================

    /// <summary>
    /// app.config配置文件帮助类
    /// </summary>
    public class AppConfigUtils
    {
        ///<summary>
        ///返回＊.exe.config文件中appSettings配置节的value项
        ///</summary>
        ///<param name="strKey">键</param>
        ///<returns>值</returns>
        public static string GetAppConfig(string strKey)
        {

            foreach (string key in ConfigurationManager.AppSettings)
            {
                if (key == strKey)
                {
                    return ConfigurationManager.AppSettings[strKey];
                }
            }
            return null;
        }

        // 更新connectionStrings配置节
        ///<summary>
        ///在＊.exe.config文件中appSettings配置节增加一对键、值对
        ///</summary>
        ///<param name="newKey">键</param>
        ///<param name="newValue">值</param>
        public static bool UpdateAppConfig(string newKey, string newValue)
        {
            foreach (string key in ConfigurationManager.AppSettings)
            {
                if (key == newKey)
                {
                    try
                    {
                        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                        config.AppSettings.Settings.Remove(newKey);
                        config.AppSettings.Settings.Add(newKey, newValue);
                        config.Save(ConfigurationSaveMode.Modified);
                        ConfigurationManager.RefreshSection("appSettings");
                        return true;
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                }
            }
            return false;
        }
    }
}
