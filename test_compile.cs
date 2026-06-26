using System.Threading.Tasks;
using Microsoft.Web.WebView2.Wpf;
class Test {
    async Task M(WebView2 w) {
        await w.EnsureCoreWebView2Async();
        w.AllowExternalDrop = false;
    }
}
