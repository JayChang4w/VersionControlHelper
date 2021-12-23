using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Extensions.Options;
using Microsoft.TeamFoundation.SourceControl.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using VersionControlHelper.Config;
using VersionControlHelper.Domain;

namespace VersionControlHelper.Service
{
    public class TfsCommitHistoryService : BaseCommitHistoryService, ITfsCommitHistoryService
    {
        private readonly TfvcHttpClient _tfvcHttpClient;

        private int? _top;
        public override int? Top
        {
            get => _top;
            set
            {
                _top = value switch
                {
                    <= 0 or > 256 or null => 100,
                    _ => value
                };
            }
        }

        public override string Project { get; set; }

        public override Enums.VersionControlType Type => Enums.VersionControlType.Tfs;

        public TfvcChangesetSearchCriteria SearchCriteria { get; set; }

        public TfsCommitHistoryService(IOptionsSnapshot<AppConfig> configSnapshot)
            : base(configSnapshot)
        {
            this._tfvcHttpClient = this._connection.GetClient<TfvcHttpClient>(); // connect to the TFS source control subpart

            this.SearchCriteria = new TfvcChangesetSearchCriteria();

            if (_config.Search.FromDate != null)
                this.SearchCriteria.FromDate = _config.Search.FromDate.Value.ToString("yyyy-MM-dd");

            if (_config.Search.ToDate != null)
                this.SearchCriteria.ToDate = _config.Search.ToDate.Value.ToString("yyyy-MM-dd");

            this.Top = _config.Search.Top;

            this.Project = _config.Project;
        }

        public override async IAsyncEnumerable<ICommitResult> GetCommitsAsync()
        {
            await foreach (var item in GetCommitsAsync(this.Project, this.SearchCriteria, this.Top))
            {
                yield return item;
            }
        }

        public async IAsyncEnumerable<ICommitResult> GetCommitsAsync(string project, TfvcChangesetSearchCriteria searchCriteria = null, int? top = null)
        {
            if (string.IsNullOrWhiteSpace(project))
                throw new ArgumentException($"'{nameof(project)}' 不得為 Null 或空白字元。", nameof(project));

            List<TfvcChangesetRef> changeSets = await _tfvcHttpClient.GetChangesetsAsync(project, searchCriteria: searchCriteria, top: top);

            foreach (TfvcChangesetRef changeSet in changeSets)
            {
                var result = new CommitResult
                {
                    Committer = changeSet.Author.DisplayName,
                    CreatedDate = changeSet.CreatedDate,
                    Comment = changeSet.Comment
                };

                List<TfvcChange> tfvcChanges = await _tfvcHttpClient.GetChangesetChangesAsync(changeSet.ChangesetId);

                result.AddChanges(tfvcChanges);

                yield return result;
            }
        }
    }
}
