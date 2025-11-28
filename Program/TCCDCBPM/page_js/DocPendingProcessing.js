$(document).ready(function () {
    getData();

    //刪除
    $(document).on("click", "a[name='delbtn']", function () {
        dguid = $(this).attr("aguid");
        if (confirm("確定要刪除該筆公文嗎？")) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/Doc_handler.aspx",
                data: {
                    crud: "d",
                    guid: dguid
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
                        alert($("Response", data).text());
                        getData();
                    }
                }
            });
        }
    });
}); // end js

function getData() {
    $.ajax({
        type: "POST",
        async: false, //在沒有返回值之前,不會執行下一步動作
        url: "../handler/Doc_handler.aspx",
        data: {
            crud: "r",
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
                        tabstr += '<td class="align-middle">' + $(this).children("建立者").text().trim() + '</td>';
                        tabstr += '<td class="align-middle">' + $(this).children("申請日期").text().trim() + '</td>';
                        tabstr += '<td class="align-middle">' + $(this).children("表單狀態_V").text().trim() + '</td>';
                        tabstr += '<td class="text-center"><div class="btn-group btn-group-sm" role="group">';
                        if ($(this).children("表單狀態_V").text().trim() == '已刪除') {
                            tabstr += '<a target="_blank" href="DocView.aspx?filetype=new&mode=view&guid='
                                + $(this).children("公文guid").text().trim() + '&parentGuid=' + $(this).children("父層guid").text().trim()
                                + '&sn=' + $(this).children("排序").text().trim() + '" class="btn btn-outline-success">檢視</a>';
                        }
                        else {
                            tabstr += '<a href="javascript:void(0);" aguid="' + $(this).children("guid").text().trim()
                                + '" name="delbtn" class="btn btn-outline-danger" name="editbtn">刪除</a>';
                            tabstr += '<a target="_blank" href="DocManageAdd.aspx?filetype=new&mode=edit&guid='
                                + $(this).children("公文guid").text().trim() + '&parentGuid=' + $(this).children("父層guid").text().trim()
                                + '&sn=' + $(this).children("排序").text().trim() + '" class="btn btn-outline-dark">編輯</a>';
                            tabstr += '<a href="DocManageAdd.aspx?filetype=demo&mode=edit&parentGuid='
                                + $(this).children("類別guid").text().trim() + '&sn=' + $(this).children("附件排序").text().trim()
                                + '" class="btn btn-outline-primary" name="editbtn">簽辦</a>';
                            tabstr += '</div></td></tr>';
                        }                        
                        Table_Select = !Table_Select;
                    });
                }
                else
                    tabstr += '<tr><td colspan="8">查詢無資料</td></tr>';

                $("#tablist tbody").append(tabstr);

                //Page.Option.Selector = "#pageblock";
                //Page.CreatePage(p, $("total", data).text());
            }
        }
    });
}