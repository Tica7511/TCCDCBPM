$(document).ready(function () {
    getData();

    // 查詢
    $(document).on("click", "#SearchBtn", function () {
        getData();
    });

    // 編輯
    //$(document).on("click", "a[name='editbtn']", function () {
    //    sessionStorage.setItem("parentguid", $(this).attr("gid"));
    //    $(this).attr("href", "fineAdd.aspx");
    //});

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
                        tabstr += '<td class="align-middle">' + $(this).children("原檔名").text().trim() + '</td>';
                        tabstr += '<td class="align-middle">' + $(this).children("lv1").text().trim() + ' > ' + $(this).children("lv2").text().trim();
                        if ($(this).children("lv3").text().trim() != '')
                            tabstr += ' > ' + $(this).children("lv3").text().trim();
                        tabstr += '</td>';
                        tabstr += '<td class="align-middle">' + $(this).children("上傳日期").text().trim() + '</td>';
                        tabstr += '<td class="text-center"><div class="btn-group btn-group-sm" role="group">';
                        tabstr += '<a href="DocManageAdd.aspx?filetype=demo&mode=edit&parentGuid='
                            + $(this).children("類別guid").text().trim() + '&sn=' + $(this).children("附件排序").text().trim()
                            + '" class="btn btn-outline-dark" name="editbtn">編輯</a>';
                        tabstr += '</div></td></tr>';
                        Table_Select = !Table_Select;
                    });
                }
                else
                    tabstr += '<tr><td colspan="5">查詢無資料</td></tr>';

                $("#tablist tbody").append(tabstr);

                //Page.Option.Selector = "#pageblock";
                //Page.CreatePage(p, $("total", data).text());
            }
        }
    });
}