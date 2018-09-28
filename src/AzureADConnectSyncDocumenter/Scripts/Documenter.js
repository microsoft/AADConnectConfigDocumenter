window.onload = function (e) {
    document.getElementById("OnlyShowChanges").disabled = false;
    document.getElementById("HideDefaultSyncRules").disabled = false;
    document.getElementById("HideEndToEndFlowsSummary").disabled = false;
}

function ToggleVisibility() {
    var x = document.getElementById("OnlyShowChanges");
    var elements = document.getElementsByClassName("CanHide");
    for (var i = 0; i < elements.length; ++i) {
        if (x.checked == true) {
            elements[i].style.display = "none";
        }
        else {
            elements[i].style.display = "";
        }
    }

    var downloadLink = document.getElementById("DownloadLink");
    if (x.checked == true) {
        downloadLink.style.display = "";
        DownloadScript(downloadLink);
    }
    else {
        downloadLink.style.display = "none";
    }
}

function ToggleDefaultRuleVisibility() {
    var x = document.getElementById("HideDefaultSyncRules");
    var elements = document.getElementsByClassName("DefaultRuleCanHide");
    for (var i = 0; i < elements.length; ++i) {
        if (x.checked == true) {
            elements[i].style.display = "none";
        }
        else {
            elements[i].style.display = "";
        }
    }
}

function ToggleEndToEndFlowsSummaryVisibility() {
    var x = document.getElementById("HideEndToEndFlowsSummary");
    var elements = document.getElementsByClassName("EndToEndFlowsSummary");
    for (var i = 0; i < elements.length; ++i) {
        if (x.checked == true) {
            elements[i].style.display = "none";
        }
        else {
            elements[i].style.display = "";
        }
    }
}

function DownloadScript(downloadLink) {
    var scripts = document.getElementsByClassName("PowerShellScript");
    var data = "";

    for (var i = 0; i < scripts.length; ++i) {
        data += scripts[i].innerText + "\r\n";
    }

    data += "\r\n#############################################################################################################################################"
    data += "\r\n### Footer Script"
    data += "\r\n#############################################################################################################################################"
    data += "\r\n"

    if (scripts.length == 1) {
        data += "\r\nWrite-Host \"There are no changes detected in the sync rule configuration of the servers.\" -ForegroundColor Green"
        data += "\r\nWrite-Host \"Please review the report manually to confim.\" -ForegroundColor Green\r\n"
    }
    else {
        data += "\r\nif ($Error.Count) {"
        data += "\r\n    Write-Host \"There were errors while executing the script.\" -ForegroundColor Red"
        data += "\r\n    Write-Host \"Please review the console output.\" -ForegroundColor Red"
        data += "\r\n}"
        data += "\r\nelse {"
        data += "\r\n    Write-Host \"Script execution completed sucessfully.\" -ForegroundColor Green"
        data += "\r\n    Write-Host \"Please regenerate the report with the latest config exports to confirm.\" -ForegroundColor Green"
        data += "\r\n    Write-Host \"Once confirmed, it is recommended that you run Full Synchronization run profile all connectors.\" -ForegroundColor Green"
        data += "\r\n}"
        data += "\r\n"
    }
    
    data = data.replace(/([^\r])\n/g, "$1\r\n");
    var file = new Blob([data], { type: "text/plain; charset=utf-8" });
    if (downloadLink == null && navigator.msSaveOrOpenBlob != null) {
        navigator.msSaveOrOpenBlob(file, "SyncRuleChanges.ps1.txt");
    }
    else {
        var downloadLink = document.getElementById("DownloadLink");
        var href = URL.createObjectURL(file);
        downloadLink.href = href;
        downloadLink.download = "SyncRuleChanges.ps1.txt"
        URL.revokeObjectURL(url);
    }

    return false;
}