import SwiftUI
import WebKit

// Replace this with your deployed Streamlit URL.
let APP_URL = "https://your-app-name.streamlit.app"

@main
struct SalaryWebViewApp: App {
    var body: some Scene {
        WindowGroup {
            WebView(urlString: APP_URL)
                .ignoresSafeArea()
        }
    }
}

struct WebView: UIViewRepresentable {
    let urlString: String

    func makeUIView(context: Context) -> WKWebView {
        let configuration = WKWebViewConfiguration()
        configuration.allowsInlineMediaPlayback = true
        let webView = WKWebView(frame: .zero, configuration: configuration)
        webView.allowsBackForwardNavigationGestures = true
        return webView
    }

    func updateUIView(_ webView: WKWebView, context: Context) {
        guard let url = URL(string: urlString) else { return }
        if webView.url == nil {
            webView.load(URLRequest(url: url))
        }
    }
}
