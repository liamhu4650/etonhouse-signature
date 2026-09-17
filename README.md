# EtonHouse Signature Studio

一个面向 Microsoft 365 Outlook 的静态电子邮件签名制作器。员工在 Outlook 中填写一次资料，配置保存到其 Exchange Online 邮箱，并在 Windows 经典版、新版 Outlook、Outlook 网页版、Outlook for Mac，以及 iOS/Android Outlook 中自动应用。

## 当前原型

- 四种学校类型：HQ、E-Bridge Pre-School、International School、EtonHouse Pre-School。
- 上传照片后自动裁剪为 256 × 256 高清头像；竖版证件照会向上取景以保留头顶，并压缩至适合 Outlook roaming settings 的大小。
- 姓名、职位、机构、电话、邮箱、网站和地址实时预览。
- 浏览器中使用本地预览存储；运行在 Outlook 中时使用 `Office.context.roamingSettings`。
- 桌面端和移动端均通过 `OnNewMessageCompose` 事件自动调用 `setSignatureAsync()`。
- 个人照片通过 CID 内嵌附件加入邮件，不需要图片服务器或数据库。
- 新模板采用左侧蓝色照片区、中部个人资料、右侧学校 Logo 和底部五品牌条。
- 学校 Logo、联系方式图标和底部品牌条均直接取自最新 PowerPoint 模板。

## 本地预览

在此目录启动任意静态 HTTP 服务器，然后打开 `index.html`。普通浏览器中不会连接 M365，保存按钮只保存本地预览。

## Outlook 测试与发布

GitHub 用户名已经配置为 `liamhu4650`。请将仓库命名为 `etonhouse-signature`；GitHub Pages 发布后，网站地址将是 `https://liamhu4650.github.io/etonhouse-signature/`。之后可在 Microsoft 365 管理中心上传 `manifest-template.xml`。

公开仓库只包含应用代码、品牌图片和明显的虚拟示例资料。员工填写的资料与个人照片不会上传到 GitHub：浏览器预览保存在该浏览器中，Outlook 配置保存在员工自己的 M365 邮箱设置中。

正式发布前仍需：

1. 如能取得原始透明 Logo，可替换当前从 JPG 模板裁切的素材以获得最佳清晰度。
2. 在 Windows 经典版、新版 Outlook、网页版、Outlook for Mac 及 iOS/Android Outlook 上完成试点测试。
3. 根据租户类型选择国际版或世纪互联版 Office.js CDN。

## 移动端部署与测试

- `manifest-template.xml` 已包含 Outlook iOS/Android 的 `MobileFormFactor`，并沿用原加载项 ID；更新清单后，现有用户分配和已保存的签名资料无需重建。
- 在 Microsoft 365 管理中心打开现有 **EtonHouse Signature**，选择“更新”，上传新版 `manifest-template.xml`。不要删除后重新部署。
- 部署同步后，请完全退出并重新打开 Outlook 手机客户端，再新建一封普通邮件进行测试。
- 移动端要求使用 Microsoft 365 或 Outlook.com 邮箱并保持联网；从 iOS 系统分享菜单创建邮件时不会触发 `OnNewMessageCompose`。
