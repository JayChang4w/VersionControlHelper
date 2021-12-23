using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.TeamFoundation.SourceControl.WebApi;
using Microsoft.VisualStudio.Services.Common;

namespace VersionControlHelper.Domain
{
    public interface ICommitResult
    {
        /// <summary>
        /// 註解
        /// </summary>
        string Comment { get; }

        /// <summary>
        /// 推送人
        /// </summary>
        string Committer { get; }

        /// <summary>
        /// 變更日期
        /// </summary>
        DateTime CreatedDate { get; }

        /// <summary>
        /// 變更項目
        /// </summary>
        ICollection<ChangeResult> Changes { get; }
    }

    public class CommitResult : ICommitResult
    {
        /// <summary>
        /// 註解
        /// </summary>
        public string Comment { get; set; }

        /// <summary>
        /// 推送人
        /// </summary>
        public string Committer { get; set; }

        /// <summary>
        /// 變更日期
        /// </summary>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// 變更項目
        /// </summary>
        public ICollection<ChangeResult> Changes { get; set; } = new List<ChangeResult>();

        /// <summary>
        /// 新增變更項目
        /// </summary>
        /// <typeparam name="TItem"></typeparam>
        /// <param name="change"></param>
        public void AddChange<TItem>(Change<TItem> change)
            where TItem : ItemModel
        {
            this.Changes.Add(new ChangeResult(change.ChangeType, change.Item));
        }

        /// <summary>
        /// 新增變更項目
        /// </summary>
        /// <typeparam name="TItem"></typeparam>
        /// <param name="changes"></param>
        public void AddChanges<TItem>(IEnumerable<Change<TItem>> changes)
            where TItem : ItemModel
        {
            this.Changes.AddRange(changes.Select(s => new ChangeResult(s.ChangeType, s.Item)));
        }
    }

    /// <summary>
    /// 變更項目
    /// </summary>
    public class ChangeResult
    {
        public ChangeResult(VersionControlChangeType changeType, ItemModel item)
        {
            ChangeType = changeType;
            Item = item;
        }

        /// <summary>
        /// 變更項目
        /// </summary>
        public ItemModel Item { get; private set; }

        /// <summary>
        /// 變更類型
        /// </summary>
        public VersionControlChangeType ChangeType { get; private set; }

        /// <summary>
        /// 變更路徑
        /// </summary>
        public string Path => Item.Path;

        /// <summary>
        /// 是否為資料夾
        /// </summary>
        public bool IsFolder => Item.IsFolder;
    }
}
