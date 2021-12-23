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
    public class GitCommitHistoryService : BaseCommitHistoryService, IGitCommitHistoryService
    {
        protected readonly GitHttpClient _gitClient;

        public override int? Top
        {
            get => SearchCriteria?.Top;
            set
            {
                SearchCriteria ??= new GitQueryCommitsCriteria();

                SearchCriteria.Top = value switch
                {
                    //<= 0 or > 256 or null => 100,
                    _ => value
                };
            }
        }

        public override string Project { get; set; }

        public override Enums.VersionControlType Type => Enums.VersionControlType.Git;

        public GitQueryCommitsCriteria SearchCriteria { get; set; }

        public GitCommitHistoryService(IOptionsSnapshot<AppConfig> configSnapshot)
            : base(configSnapshot)
        {
            this._gitClient = this._connection.GetClient<GitHttpClient>();

            this.SearchCriteria = new GitQueryCommitsCriteria()
            {
                ItemVersion = new GitVersionDescriptor()
                {
                    Version = string.IsNullOrWhiteSpace(_config.GitBranch) ? "master" : _config.GitBranch,
                    VersionType = GitVersionType.Branch
                }
            };

            if (_config.Search.FromDate != null)
                this.SearchCriteria.FromDate = _config.Search.FromDate.Value.ToString("yyyy-MM-dd");

            if (_config.Search.ToDate != null)
                this.SearchCriteria.ToDate = _config.Search.ToDate.Value.ToString("yyyy-MM-dd");

            this.Top = _config.Search.Top;

            this.Project = _config.Project;
        }

        public override async IAsyncEnumerable<ICommitResult> GetCommitsAsync()
        {
            await foreach (var item in GetCommitsAsync(this.Project, this.SearchCriteria))
            {
                yield return item;
            }
        }

        public async IAsyncEnumerable<ICommitResult> GetCommitsAsync(string project, GitQueryCommitsCriteria searchCriteria = null)
        {
            if (string.IsNullOrWhiteSpace(project))
                throw new ArgumentException($"'{nameof(project)}' 不得為 Null 或空白字元。", nameof(project));

            List<GitRepository> repositories = await _gitClient.GetRepositoriesAsync();

            GitRepository repos = repositories.Where(w => w.Name == project).SingleOrDefault() ?? throw new ConfigFileException("Project Not Found");

            List<GitCommitRef> commitRefs = await _gitClient.GetCommitsAsync(repos.Id, searchCriteria);

            foreach (GitCommitRef commitRef in commitRefs)
            {
                var result = new CommitResult
                {
                    Committer = commitRef.Committer.Name,
                    CreatedDate = commitRef.Committer.Date,
                    Comment = commitRef.Comment
                };

                GitCommitChanges gitCommitChanges = await _gitClient.GetChangesAsync(commitRef.CommitId, repos.Id.ToString());

                result.AddChanges(gitCommitChanges.Changes);

                yield return result;
            }
        }
    }
}
