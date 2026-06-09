# iOS WebView Shell

這不是完整 `.ipa`，而是 SwiftUI + WKWebView 的外殼程式碼。

使用方式：

1. 用 Xcode 建立新 iOS App 專案。
2. 把 `SalaryWebViewApp.swift` 的內容貼進專案。
3. 將 `APP_URL` 改成你部署好的 Streamlit 網址。
4. 設定 Signing & Capabilities。
5. 用 Xcode Build 到你的 iPhone，或透過 App Store Connect 發布。

注意：Apple 要求 iOS App 必須簽名；沒有 Apple Developer 簽名流程，無法產出可給別人直接安裝的正式 `.ipa`。
