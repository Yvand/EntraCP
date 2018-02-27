/// <reference path="../node_modules/@types/jquery/index.d.ts" />
// npm install --save @types/jquery

namespace GitHubStatistics {    
    export interface RepoStatsJSON {
        DateStatCreatedUTC: string;
        Repository: string;
        LatestAssetUrl: string;
        LatestReleaseCreationDate: string;
        LatestReleaseTagName: string;
        LatestReleaseDownloadCount: number;
        AllReleasesDownloadCount: number;
        TotalDownloadCount: number;
    }

    export class RepoStats {
        url: string = "http://ldapcp-functions.azurewebsites.net/api/GetLatestAzureCPRepoStats";
        //url: string = "http://jsfiddle.net/echo/jsonp/";
        authZKey: string = "3YWIoBPB2gMnG4bN5WsTF4p14V3Bx7U7ZXqrRX2SWj13Mn2omw3OnQ==";
        getLatestStat() {
            //console.log("Sending query to " + this.url);            
            $.ajax({
                method: "GET",
                crossDomain: true,
                data: {code: this.authZKey},
                dataType: "jsonp",
                jsonpCallback: "GitHubStatistics.RepoStats.parseGitHubStatisticsResponse",
                url: this.url,
                success: function(responseData, textStatus, jqXHR) {
                },
                error: function (responseData, textStatus, errorThrown) {
                    console.log("Request to " + this.url + " failed: " + errorThrown);
                }
            });
        }

        static decodeJSONResponse(json: GitHubStatistics.RepoStatsJSON) {
            var obj = Object.assign({}, json, {
                //created: new Date(json.DateStatCreatedUTC)
            });
            return obj;
        }

        static parseGitHubStatisticsResponse (data) {
            var result =  GitHubStatistics.RepoStats.decodeJSONResponse(data);
            $("#TotalDownloadCount").text(result.TotalDownloadCount);
            $("#LatestReleaseDownloadCount").text(result.LatestReleaseDownloadCount);
            $("#LatestReleaseTagName").text(result.LatestReleaseTagName);
            $("#LatestAssetUrl").attr("href", result.LatestAssetUrl)
            //$("#LatestReleaseCreationDate").text(result.LatestReleaseCreationDate);
        };
    }
}

$(document).ready(function () {
    let stats = new GitHubStatistics.RepoStats();
    let result = stats.getLatestStat()
});

