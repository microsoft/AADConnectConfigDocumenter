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

function DownloadScript(downloadLink) {
    var scripts = document.getElementsByClassName("PowerShellScript");
    var data = "";
    for (var i = 0; i < scripts.length; ++i) {
        data += scripts[i].innerText;
    }

    var file = new Blob([data.replace(/([^\r])\n/g, "$1\r\n")], { type: "text/plain; charset=utf-8" });
    if (downloadLink == null && navigator.msSaveOrOpenBlob != null) {
        navigator.msSaveOrOpenBlob(file, "SyncRuleChanges.ps1.txt");
        return false;
    }

    var downloadLink = document.getElementById("DownloadLink");
    var href = URL.createObjectURL(file);
    downloadLink.href = href;
    downloadLink.download = "SyncRuleChanges.ps1.txt"
}