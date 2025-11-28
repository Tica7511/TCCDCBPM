$(document).ready(function () {
    getData();

    // 查詢
    $(document).on("click", "#SearchBtn", function () {
        getData();
    });

    //編輯
    $(document).on("click", "a[name='searchbtn']", function () {
        $("#mp_filename").html($(this).attr("afilename"));
        $("#h_Guid").val($(this).attr("aGuid"));
        $("#h_parentGuid").val($(this).attr("aparentGuid"));
        $("#h_version").val($(this).attr("aversion"));
        $("#h_nowVersion").val($(this).attr("aversion"));
        getVersionList($("#h_Guid").val());
        $("#mgsellist").val($("#h_nowVersion").val());
        getVersionData($("#h_Guid").val(), $("#h_nowVersion").val());
        doOpenMagPopup();
    });

    //選擇版本
    $(document).on("change", "#mgsellist", function () {
        getVersionData($("#h_Guid").val() ,$("#mgsellist option:selected").val());
    });

    //項目名稱 編輯按鈕
    $(document).on("click", "a[name='a_editbtn']", function () {
        var atypecode = $(this).attr("atypecode");
        var atypename = $("span[name='sp_typecode_" + atypecode + "']").html();
        $("#a_editbtn_" + atypecode).hide();
        $("span[name='sp_typecode_" + atypecode + "']").hide();
        $("input[name='txt_typecode_" + atypecode + "']").show();
        $("input[name='txt_typecode_" + atypecode + "']").val(atypename);
        $("#a_delbtn_" + atypecode).show();
        $("#a_save_" + atypecode).show();
    });

    //項目名稱 返回按鈕
    $(document).on("click", "a[name='a_delbtn']", function () {
        var atypecode = $(this).attr("atypecode");
        $("#a_editbtn_" + atypecode).show();
        $("span[name='sp_typecode_" + atypecode + "']").show();
        $("input[name='txt_typecode_" + atypecode + "']").hide();
        $("#a_delbtn_" + atypecode).hide();
        $("#a_save_" + atypecode).hide();
    });

    //項目名稱 儲存按鈕
    $(document).on("click", "a[name='a_save']", function () {
        var fguid = $("#h_Guid").val();
        var atypecode = $(this).attr("atypecode");
        var atypename = $("input[name='txt_typecode_" + atypecode + "']").val();

        console.log("fguid:", fguid, "atypecode:", atypecode, "atypename:", atypename);

        if (confirm("確定儲存嗎?")) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/DocFieldManage_handler.aspx",
                data: {
                    fguid: fguid,
                    atypecode: atypecode,
                    atypename: atypename,
                    crud: "cu",
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
                        $("span[name='sp_typecode_" + atypecode + "']").html($("input[name='txt_typecode_" + atypecode + "']").val());
                        $("#a_editbtn_" + atypecode).show();
                        $("span[name='sp_typecode_" + atypecode + "']").show();
                        $("input[name='txt_typecode_" + atypecode + "']").hide();
                        $("#a_delbtn_" + atypecode).hide();
                        $("#a_save_" + atypecode).hide();
                        alert("儲存成功!");
                    }
                }
            });
        }
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
        url: "../handler/Doc_handler.aspx",
        data: {
            crud: "r",
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
                        tabstr += '<td class="align-middle">' + $(this).children("表單名稱").text().trim() + '</td>';
                        tabstr += '<td class="align-middle">' + $(this).children("表單編號").text().trim() + '</td>';
                        tabstr += '<td class="align-middle">' + $(this).children("主旨").text().trim() + '</td>';
                        tabstr += '<td class="align-middle">' + $(this).children("上傳日期").text().trim() + '</td>';
                        tabstr += '<td class="align-middle">' + $(this).children("表單狀態_V").text().trim() + '</td>';
                        tabstr += '<td class="align-middle text-center"><div class="btn-group btn-group-sm" role="group">';
                        tabstr += '<a href="javascript:void(0);" name="searchbtn" aparentGuid="' + $(this).children("類別guid").text().trim()
                            + '" aversion="' + $(this).children("版本").text().trim() + '" afilename="' + $(this).children("表單名稱").text().trim()
                            + '-' + $(this).children("主旨").text().trim() + '" aGuid="' + $(this).children("公文guid").text().trim()
                            + '" class="btn btn-outline-dark">查詢</a> ';
                        tabstr += '</div></td>';
                        tabstr += '<td class="align-middle text-center"><div class="btn-group btn-group-sm" role="group">';
                        tabstr += '<a target="_blank" href="DocManageAdd.aspx?filetype=new&mode=view&guid=' + $(this).children("公文guid").text().trim() + '&parentGuid='
                            + $(this).children("父層guid").text().trim() + '&sn=' + $(this).children("排序").text().trim()
                            + '" class="btn btn-outline-success">預覽</a>';
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

function getVersionData(guid, version) {
    $.ajax({
        type: "POST",
        async: false, //在沒有返回值之前,不會執行下一步動作
        url: "../handler/Doc_handler.aspx",
        data: {
            crud: "rvd",
            guid: guid,
            version: version
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
                $("#tablistFileSearch tbody").empty();
                var tabstr = '';
                if ($(data).find("data_item").length > 0) {
                    var Table_Select = false;
                    $(data).find("data_item").each(function (i) {
                        $("#h_version").val($(this).children("版本").text().trim());
                        var typecode = $(this).children("項目代碼").text().trim();
                        var typeName = $(this).children("項目名稱").text().trim();
                        var version = $(this).children("版本").text().trim();
                        tabstr += Table_Select ? '<tr class="Table_Select">' : '<tr>';
                        tabstr += '<td>' + $(this).children("項目代碼").text().trim() + '</td>';
                        //tabstr += '<td>';
                        //if (version == $("#h_nowVersion").val()) {
                        //    tabstr += '<span name="sp_typecode_' + typecode + '">' + typeName + '</span>'
                        //        + '<input style="display:none;" type="text" class="inputex width60" name="txt_typecode_' + typecode + '" value="' + typeName + '"/>'
                        //        + ' <a id="a_editbtn_' + typecode + '" name="a_editbtn" atypecode="' + typecode + '" href="javascript:void(0);">編輯</a>'
                        //        + '<a id="a_delbtn_' + typecode + '" style="display:none;" href="javascript:void(0);" atypecode="' + typecode + '" name="a_delbtn">返回</a> '
                        //        + '<a id="a_save_' + typecode + '" style="display:none;" href="javascript:void(0);" atypecode="' + typecode + '" name="a_save">儲存</a>';
                        //}
                        //else {
                        //    tabstr += $(this).children("項目名稱").text().trim();
                        //}
                        //tabstr += '</td>';
                        tabstr += '<td>' + $(this).children("項目名稱").text().trim() + '</td>';
                        tabstr += '<td>' + $(this).children("項目內容").text().trim() + '</td>';
                        tabstr += '</tr>';
                        Table_Select = !Table_Select;
                    });
                }
                else
                    tabstr += '<tr><td colspan="3">目前版本尚無任何內容控制項</td></tr>';

                $("#tablistFileSearch tbody").append(tabstr);

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
        url: "../handler/Doc_handler.aspx",
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
                        ddlstr += '<option value="' + $(this).children("版本").text().trim() + '">第' + $(this).children("版本").text().trim() + '版</option>'
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
            src: '#mp_filesearch'
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