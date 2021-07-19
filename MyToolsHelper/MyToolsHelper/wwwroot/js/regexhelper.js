//获取url参数值
function getQueryString(url, name) {
    var reg = new RegExp("(?<=[?&]{1}" + name + "=)[^&]*");
    var r = url.match(reg);
    if (r != null) {
        return unescape(r[0]);
    } else {
        return null;
    }
}