$(document).ready(function () {
    getData();
});

//取得資料庫檔案
function getData() {
    var crud = $.getQueryString("filetype");
    var parentGuid = $.getQueryString("parentGuid");
    var sn = $.getQueryString("sn");
    var mode = $.getQueryString("mode");
    var status = true;
    if (mode == 'view')
        status = false; // 如果是查看模式，則不允許編輯

    $.ajax({
        type: "POST",
        async: true,
        url: "../handler/file_handler.aspx",
        data: {
            crud: crud, 
            parentGuid: parentGuid,
            sn: sn,
            mode: mode
        },
        error: function (xhr) {
            alert("Error: " + xhr.status);
            console.log(xhr.responseText);
        },
        success: function (data) {

            if ($(data).find("Error").length > 0) {
                alert($(data).find("Error").attr("Message"));
            }
            else {
                if ($(data).find("data_item").length > 0) {
                    $(data).find("data_item").each(function (i) {
                        console.log("後端回傳檔案名：", $("fileName", data).text());
                        const fileName = $("fileName", data).text();
                        const fileNewName = $("fileNewName", data).text();
                        const authToken = $("token", data).text();
                        const onlyofficeguid = $("onlyofficeguid", data).text();
                        const parentguid = $("parentguid", data).text();
                        const version = $("version", data).text();
                        const mGuid = $("mGuid", data).text();
                        const mName = $("mName", data).text();
                        $("#fGuid").val(onlyofficeguid);
                        localStorage.setItem('authToken', authToken);

                        console.log("Initializing OnlyOffice Viewer for:", fileName, "\rToken:", authToken);

                        var editorConfig = {
                            "document": {
                                "fileType": "docx",
                                "key": onlyofficeguid + '_' + parentguid + '_' + version, //附件guid + 父層guid + 版本
                                "title": fileName,
                                "url": "http://172.20.10.5:7594/DOWNLOAD.aspx?category=Demo&guid=" + onlyofficeguid
                            },
                            "documentType": "word",
                            "editorConfig": {
                                "mode": mode,
                                "lang": "zh-TW",
                                "canUseHistory": true,
                                "callbackUrl": "http://172.20.10.5:7594/handler/SaveCallback.aspx",
                                "customization": {
                                    "forcesave": true,
                                    "autosave": true,
                                    "autosaveInterval": 60,
                                    /*"trackChanges": true,*/
                                    "buttons": {
                                        "print": false,
                                        "download": false
                                    },
                                    "logo": {
                                        image: "http://172.20.10.5:7594/images/tccLogo.png",
                                        url: "https://www.cogen.com.tw/tw/"
                                    },
                                    "layout": {
                                        "leftMenu": { mode: false },  // 左側欄初始隱藏
                                        "rightMenu": { mode: false },   // 右側欄初始隱藏
                                        "toolbar": {
                                            "layout": false, // 版面配置
                                            "references": false, // 參考文獻
                                            "protect": false, // 保護
                                            //"plugins": false  // 外掛程式
                                            // 其他可控分頁（示意）：file, home, insert, layout, draw, view, collaboration ...
                                        }
                                    },
                                },
                                "user": {
                                    "id": mGuid,
                                    "name": mName
                                },
                                "history": {
                                    "serverVersion": true
                                }
                            },
                            "permissions": {
                                "edit": status,
                                /*"review": true,*/  // 開啟追蹤修訂按鈕
                                "comment": status,
                                "print": false,
                                "download": false
                            },
                            "token": authToken,
                            "height": "800px",
                            "width": "100%",
                            "type": "desktop",
                            "events": {
                                onRequestHistory: handleRequestHistory,
                                onRequestHistoryData: handleRequestHistoryData,
                                onRequestHistoryClose: handleRequestHistoryClose,
                                onRequestRestore: handleRequestRestore
                            }
                        };

                        lastEditorConfig = editorConfig;
                        docEditor = new DocsAPI.DocEditor("div_container", editorConfig);

                        //setInterval(function () {
                        //    $.post("../Handler/TriggerForceSave.aspx", {
                        //        key: onlyofficeguid
                        //    });
                        //}, 60000); 
                    });
                }
                else {
                    alert("nothing");
                }
            }
        }
    });
}

function initEditor(config) {
    if (docEditor && typeof docEditor.destroyEditor === "function") {
        docEditor.destroyEditor(); // 清掉原本的 editor
    }
    docEditor = new DocsAPI.DocEditor("div_container", config);
}

function handleRequestHistory() {
    console.log("onRequestHistory triggered");

    // 確保 docEditor 已經完成初始化
    if (!docEditor || typeof docEditor.refreshHistory !== "function") {
        console.warn("docEditor not ready");
        return;
    }

    const onlyofficeguid = $("#fGuid").val();

    $.ajax({
        type: "GET",
        url: "http://172.20.10.5:7594/Handler/DocumentHistoryHandler.aspx",
        data: {
            type: "01",
            guid: onlyofficeguid
        },
        dataType: "json",
        success: function (historyResponse) {
            console.log("Received history response:", historyResponse);
            docEditor.refreshHistory(historyResponse);
            console.log("refreshHistory called with:", historyResponse);
        },
        error: function (xhr) {
            console.error("Failed to get history data:", xhr);
        }
    });
}

function handleRequestHistoryData(event) {
    console.log("onRequestHistoryData event triggered. Fetching history...");

    // 確保 docEditor 已初始化
    if (!docEditor || typeof docEditor.refreshHistory !== "function") {
        console.warn("docEditor not ready");
        return;
    }

    const version = event.data;
    const onlyofficeguid = $("#fGuid").val();

    $.ajax({
        type: "GET",
        url: "http://172.20.10.5:7594/Handler/DocumentHistoryHandler.aspx",
        data: {
            type: "02",
            guid: onlyofficeguid,
            version: version
        },
        dataType: "json",
        success: function (historyResponse) {
            console.log("Received history response:", historyResponse);

            if (docEditor && typeof docEditor.refreshHistory === "function") {
                console.log("Calling docEditor.setHistoryData");
                docEditor.setHistoryData(historyResponse);
                console.log("complete docEditor.setHistoryData");
            } else {
                console.error("setHistoryData is not a function or docEditor not ready");
            }
        },
        error: function (xhr) {
            console.error("Failed to get history data:", xhr);
        }
    });
}

function handleRequestHistoryClose() {
    document.location.reload();
    //console.log("關閉歷史檢視，將重新啟動編輯器");
    //if (lastEditorConfig) {
    //    initEditor(lastEditorConfig); // 重啟編輯器，不 reload
    //} else {
    //    console.warn("找不到之前的 editor config");
    //}
}

function handleRequestRestore(event) {
    const version = event.data.version;
    const fileUrl = event.data.url;
    const fileType = event.data.fileType;
    const changes = event.data.changes;
    const onlyofficeguid = $("#fGuid").val();

    console.log("guid:", $("#fGuid").val());
    console.log("使用者請求還原版本:", version);
    console.log("還原檔案連結:", fileUrl);
    console.log("變更摘要:", changes);

    $.ajax({
        type: "GET",
        url: "http://172.20.10.5:7594/Handler/DocumentHistoryHandler.aspx",
        data: {
            type: "03",
            guid: onlyofficeguid,
            version: version
        },
        dataType: "json",
        success: function () {
            alert("成功還原版本，重新載入最新內容");
            document.location.reload();
        },
        error: function (xhr) {
            console.error("Failed to get history data:", xhr);
        }
    });
}