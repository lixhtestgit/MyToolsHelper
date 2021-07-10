function PayFetchAjax() {
    var _this = this;
    var ajaxUrl = "";
    var data = {};
    var postContentType = "application/json";
    var rollBack = function (result) {
    };

    //请求地址，传输对象：Object类型，回调方法
    _this.Init = function (actionUrl, dataObject, rollBackFunc) {
        ajaxUrl = actionUrl;
        data = dataObject;
        rollBack = rollBackFunc;
        return _this;
    };

    _this.SetFormContentType = function () {
        postContentType = "application/x-www-form-urlencoded";
        var formData = [];
        for (var pro in data) {
            formData.push(pro + "=" + window.encodeURIComponent(data[pro]));
        }
        data = formData.join("&");
        return _this;
    };

    //get
    _this.Get = function () {
        return fetch(ajaxUrl, {
            method: 'GET',
            credentials: 'include'
        }).then(function (response) {
            return response.json();
        }).then(rollBack);
    };

    //post
    _this.Post = function () {

        var bodyValue = "";
        if (postContentType.indexOf("json") > -1) {
            bodyValue = JSON.stringify(data);
        } else if (postContentType.indexOf("form") > -1) {
            bodyValue = data;
        }

        return fetch(ajaxUrl, {
            method: 'POST',
            credentials: 'include',
            mode: "cors",
            headers: {
                'Content-Type': postContentType
            },
            body: bodyValue
        }).then(function (response) {
            return response.json();
        }).then(rollBack);
    };
}