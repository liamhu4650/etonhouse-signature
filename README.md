# EtonHouse Signature Studio

一个面向 Microsoft 365 Outlook 的静态电子邮件签名制作器。员工在 Outlook 中填写一次资料，配置保存到其 Exchange Online 邮箱，并在 Windows 经典版、新版 Outlook、Outlook 网页版和 Outlook for Mac 中自动应用。

## 当前原型

- 四种学校类型：HQ、E-Bridge Pre-School、International School、EtonHouse Pre-School。
- 上传照片后自动居中裁剪为 128 × 128，并压缩至适合 Outlook roaming settings 的大小。
- 姓名、职位、机构、电话、邮箱、网站和地址实时预览。
- 浏览器中使用本地预览存储；运行在 Outlook 中时使用 `Office.context.roamingSettings`。
- `OnNewMessageCompose` 事件自动调用 `setSignatureAsync()`。
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
2. 在 Windows 经典版、新版 Outlook、网页版及 Outlook for Mac 上完成试点测试。
3. 根据租户类型选择国际版或世纪互联版 Office.js CDN。
