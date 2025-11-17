$(document).ready(function () {
    getData();

    // 查詢
    $(document).on("click", "#SearchBtn", function () {
        getData();
    });

    //編輯
    $(document).on("click", "a[name='editbtn']", function () {
        $("#mp_filename").html($(this).attr("afilename"));
        $("#h_Guid").val($(this).attr("aGuid"));
        $("#h_parentGuid").val($(this).attr("aparentGuid"));
        $("#h_sn").val($(this).attr("asn"));
        getVersionList($("#h_Guid").val());
        $("#mgsellist").val($("#h_Guid").val());
        getVersionData($("#h_Guid").val());
        doOpenMagPopup();
    });

    //選擇版本
    $(document).on("change", "#mgsellist", function () {
        getVersionData($("#mgsellist option:selected").val());
    });

    // 檢視
    //$(document).on("click", "a[name='viewbtn']", function () {
    //    sessionStorage.setItem("parentguid", $(this).attr("gid"));
    //    $(this).attr("href", "fineView.aspx");
    //});
}); // end js

function getData() {
    $.ajax({
        type: "POST",
        async: false, //在沒有返回值之前,不會執行下一步動作
        url: "../handler/DocMange_handler.aspx",
        data: {
            crud: "rf",
            //userid: $("#ContentPlaceHolder1_userid").val(),
            //useridx: $("#ContentPlaceHolder1_useridx").val(),
            //PageNo: p,
            //PageSize: Page.Option.PageSize,
            //whereColumn: "fm_formNumber",
            //SearchStr: $("#SearchStr").val()
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
                $("#tablist tbody").empty();
                var tabstr = '';
                if ($(data).find("data_item").length > 0) {
                    var Table_Select = false;
                    $(data).find("data_item").each(function (i) {
                        tabstr += Table_Select ? '<tr class="Table_Select">' : '<tr>';
                        tabstr += '<td class="align-middle">' + $(this).children("範本名稱").text().trim() + '</td>';
                        tabstr += '<td class="align-middle text-center">' + $(this).children("目前版本").text().trim() + '</td>';
                        tabstr += '<td class="align-middle">' + $(this).children("lv1").text().trim() + ' > ' + $(this).children("lv2").text().trim();
                        if ($(this).children("lv3").text().trim() != '')
                            tabstr += ' > ' + $(this).children("lv3").text().trim();
                        tabstr += '</td>';
                        tabstr += '<td class="align-middle text-center">' + $(this).children("欄位數").text().trim() + '</td>';
                        tabstr += '<td class="align-middle text-center">' + $(this).children("最後修改者").text().trim() + '</td>';
                        tabstr += '<td class="align-middle">' + $(this).children("最後修改時間").text().trim() + '</td>';
                        tabstr += '<td class="text-center"><div class="btn-group btn-group-sm" role="group">';
                        tabstr += '<a href="javascript:void(0);" name="editbtn" aparentGuid="' + $(this).children("類別guid").text().trim()
                            + '" asn="' + $(this).children("附件排序").text().trim() + '" afilename="' + $(this).children("範本名稱").text().trim()
                            + '" aGuid="' + $(this).children("附件guid").text().trim() + '" class="btn btn-outline-dark">編輯</a> ';
                        tabstr += '<a target="_blank" href="DocManageAdd.aspx?filetype=demo&mode=view&parentGuid='
                            + $(this).children("類別guid").text().trim() + '&sn=' + $(this).children("附件排序").text().trim()
                            + '" class="btn btn-outline-dark">預覽</a>';
                        tabstr += '</div></td></tr>';
                        Table_Select = !Table_Select;
                    });
                }
                else
                    tabstr += '<tr><td colspan="7">查詢無資料</td></tr>';

                $("#tablist tbody").append(tabstr);

                //Page.Option.Selector = "#pageblock";
                //Page.CreatePage(p, $("total", data).text());
            }
        }
    });
}

function getVersionData(guid) {
    $.ajax({
        type: "POST",
        async: false, //在沒有返回值之前,不會執行下一步動作
        url: "../handler/DocFieldManage_handler.aspx",
        data: {
            crud: "r",
            guid: guid,
            //userid: $("#ContentPlaceHolder1_userid").val(),
            //useridx: $("#ContentPlaceHolder1_useridx").val(),
            //PageNo: p,
            //PageSize: Page.Option.PageSize,
            //whereColumn: "fm_formNumber",
            //SearchStr: $("#SearchStr").val()
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
                $("#tablistField tbody").empty();
                var tabstr = '';
                if ($(data).find("data_item").length > 0) {
                    alert($(data).find("data_item"));
                    $(data).find("data_item").each(function (i) {
                        tabstr += Table_Select ? '<tr class="Table_Select">' : '<tr>';
                        tabstr += '<td class="text-center">' + $(this).children("項目代碼").text().trim() + '</td>';
                        tabstr += '<td class="text-center">' + $(this).children("項目名稱").text().trim() + '</td>';
                        tabstr += '<td class="text-center">' + $(this).children("項目類型").text().trim() + '</td>';                        
                        tabstr += '<td class="text-center">' + $(this).children("是否已刪除_V").text().trim() + '</td>';
                        tabstr += '</tr>';
                        Table_Select = !Table_Select;
                    });
                }
                else
                    tabstr += '<tr><td colspan="4">目前版本尚無任何內容控制項</td></tr>';

                $("#tablistField tbody").append(tabstr);

                //Page.Option.Selector = "#pageblock";
                //Page.CreatePage(p, $("total", data).text());
            }
        }
    });
}

function getVersionList(guid) {
    $.ajax({
        type: "POST",
        async: false, //在沒有返回值之前,不會執行下一步動作
        url: "../handler/DocMange_handler.aspx",
        data: {
            crud: "rv",
            guid: guid
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
                $("#mgsellist").empty();
                var ddlstr = '';
                if ($(data).find("data_item").length > 0) {
                    $(data).find("data_item").each(function (i) {
                        ddlstr += '<option value="' + $(this).children("guid").text().trim() + '">第' + $(this).children("版本").text().trim() + '版</option>'
                    });
                }
                $("#mgsellist").append(ddlstr);
            }
        }
    });
}

function doOpenMagPopup() {
    $.magnificPopup.open({
        items: {
            src: '#mp_field'
        },
        type: 'inline',
        midClick: false, // 是否使用滑鼠中鍵
        closeOnBgClick: true,//點擊背景關閉視窗
        showCloseBtn: true,//隱藏關閉按鈕
        fixedContentPos: true,//彈出視窗是否固定在畫面上
        mainClass: 'mfp-fade',//加入CSS淡入淡出效果
        tClose: '關閉',//翻譯字串
    });
}