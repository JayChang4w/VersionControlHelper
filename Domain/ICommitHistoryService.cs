using Microsoft.TeamFoundation.SourceControl.WebApi;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VersionControlHelper.Domain
{
    public interface ICommitHistoryService
    {
        Enums.VersionControlType Type { get; }
        int? Top { get; set; }
        string Project { get; set; }
        IAsyncEnumerable<ICommitResult> GetCommitsAsync();
        Task ExportExcelAsync(string folderPath = null, string fileName = null);
    }

    public interface IGitCommitHistoryService : ICommitHistoryService
    {
        GitQueryCommitsCriteria SearchCriteria { get; set; }
        IAsyncEnumerable<ICommitResult> GetCommitsAsync(string project, GitQueryCommitsCriteria searchCriteria = null);
    }

    public interface ITfsCommitHistoryService : ICommitHistoryService
    {
        TfvcChangesetSearchCriteria SearchCriteria { get; set; }
        IAsyncEnumerable<ICommitResult> GetCommitsAsync(string project, TfvcChangesetSearchCriteria searchCriteria = null, int? top = null);
    }
}
