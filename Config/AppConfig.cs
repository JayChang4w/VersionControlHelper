
using VersionControlHelper.Domain;

namespace VersionControlHelper.Config
{
    public class AppConfig
    {
        /// <summary>
        /// 版控類型 (Tfs or Git)
        /// </summary>
        public Enums.VersionControlType Type { get; set; }

        /// <summary>
        /// 匯出路徑
        /// </summary>
        public string ExportPath { get; set; }

        /// <summary>
        /// Server網址
        /// </summary>
        public string ServerUrl { get; set; }

        /// <summary>
        /// 專案名稱
        /// </summary>
        public string Project { get; set; }

        /// <summary>
        /// Git 分支名稱
        /// </summary>
        public string GitBranch { get; set; }

        /// <summary>
        /// 授權
        /// </summary>
        public CredentialConfig Credential { get; set; }

        /// <summary>
        /// 搜尋條件
        /// </summary>
        public SearchConfig Search { get; set; }
    }
}
