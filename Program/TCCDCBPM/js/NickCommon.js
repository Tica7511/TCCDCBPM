/*-------------------------------------
Title：Common
Author：Nick Lai
Create Date：2018/11/08
Modify Date：2018/11/08
Version：1.0
-------------------------------------*/

$.extend({
    getQueryString: function (str) {
        ///---------------------------------------------------------------------------
        /// 功    能: 抓取Get參數之變數值
        ///---------------------------------------------------------------------------
        parName = str.replace(/[\[]/, '\\\[').replace(/[\]]/, '\\\]');
        var pattern = '[\\?&]' + str + '=([^&#]*)';
        var regex = new RegExp(pattern);
        var matches = regex.exec(window.location.href);
        if (matches == null) return '';
        else return decodeURIComponent(matches[1].replace(/\+/g, ' '));
    },

    newGuid: function () {
        var guid = "";
        for (var i = 1; i <= 32; i++) {
            var n = Math.floor(Math.random() * 16.0).toString(16);
            guid += n;
            //if ((i == 8) || (i == 12) || (i == 16) || (i == 20))
            //    guid += "-";
        }
        return guid;
    },

    // 千分位
    FormatThousandGroup: function (n) {
        if (!$.isNumeric(n)) // 非數字回傳原值
            return n;
        else {
            n = Number(n); // 去小數點為0
            n += ""; // 轉字串
            var arr = n.split(".");
            var re = /(\d{1,3})(?=(\d{3})+$)/g;
            return arr[0].replace(re, "$1,") + (arr.length == 2 ? "." + arr[1] : "");
        }
    },

     // Get 時間
    FormatTime: function (t) {
        var d = new Date(t); // for now

        var h = d.getHours();
        h = (h < 10) ? ("0" + h) : h;

        var m = d.getMinutes();
        m = (m < 10) ? ("0" + m) : m;

        var s = d.getSeconds();
        s = (s < 10) ? ("0" + s) : s;

        timestr = h + ":" + m + ":" + s;
        return timestr;
    }
});

var DOMPurify = {
    sanitize: function (data) {
        return data;
    }
}