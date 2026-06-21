# macOS 签名发布配置

要让其他 Mac 正常双击运行，应用必须使用 Apple Developer Program 的 `Developer ID Application` 证书签名并提交 Apple 公证。

仓库中的 `.github/workflows/release-macos-signed.yml` 已包含构建、签名、公证、装订公证票据和 Gatekeeper 校验。

## GitHub Secrets

在仓库 `Settings > Secrets and variables > Actions` 中配置：

- `APPLE_CERTIFICATE_BASE64`：Developer ID Application `.p12` 文件的 Base64 内容。
- `APPLE_CERTIFICATE_PASSWORD`：导出 `.p12` 时设置的密码。
- `APPLE_SIGNING_IDENTITY`：例如 `Developer ID Application: Company Name (TEAMID)`。
- `KEYCHAIN_PASSWORD`：CI 临时钥匙串密码，可自行生成强密码。
- `APPLE_API_KEY_BASE64`：App Store Connect API `.p8` 文件的 Base64 内容。
- `APPLE_API_KEY_ID`：API Key ID。
- `APPLE_API_ISSUER`：App Store Connect Issuer ID。

配置后，在 GitHub Actions 中手工运行 `Release signed macOS app`，下载 `供需协同工具V3.0-macOS-signed-notarized`。

不要把 `.p12`、`.p8`、密码或数据库连接串提交到仓库。
