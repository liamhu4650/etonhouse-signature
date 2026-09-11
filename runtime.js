/* Self-contained because classic Outlook loads one event-runtime script. */
(function () {
  const profileKey = "etonhouse.signature.profile.v1";
  const productionAssetFallback = "https://liamhu4650.github.io/etonhouse-signature/assets/";
  const schools = {
    hq: { label: "HQ 总部", logoAsset: "school-hq.png" },
    ebridge: { label: "E-Bridge Pre-School", logoAsset: "school-ebridge.png" },
    international: { label: "International School", logoAsset: "school-international.png" },
    preschool: { label: "EtonHouse Pre-School", logoAsset: "school-preschool.png" },
  };

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function normalizedUrl(value) {
    const clean = String(value || "").trim();
    return !clean || /^https?:\/\//i.test(clean) ? clean : `https://${clean}`;
  }

  function assetUrl(base, fileName) {
    return `${String(base || productionAssetFallback).replace(/\/?$/, "/")}${fileName}`;
  }

  function infoRow(iconAsset, value, href, assetBase) {
    if (!value) return "";
    const content = href
      ? `<a href="${escapeHtml(href)}" style="color:#6b6b6b;text-decoration:none">${escapeHtml(value)}</a>`
      : escapeHtml(value);
    return `<tr><td width="27" style="width:27px;padding:2px 0;vertical-align:top"><img src="${escapeHtml(assetUrl(assetBase, iconAsset))}" width="20" alt="" style="display:block;width:20px;height:auto;border:0" /></td><td style="padding:2px 0;color:#6b6b6b;font-family:Arial,sans-serif;font-size:12px;font-weight:700;line-height:20px;vertical-align:top">${content}</td></tr>`;
  }

  function buildSignatureHtml(data, photoFileName) {
    const school = schools[data.school] || schools.hq;
    const assetBase = data.assetBaseUrl || productionAssetFallback;
    const changePhotoLabel = data.uiLanguage === "en" ? "CHANGE PHOTO" : "更换照片";
    const photo = photoFileName
      ? `<img src="cid:${escapeHtml(photoFileName)}" width="118" height="118" alt="${escapeHtml(data.name)}" style="display:block;width:118px;height:118px;border:0" />`
      : `<div style="display:table-cell;width:118px;height:118px;color:#fff;background:#d71920;font-family:Microsoft YaHei,Arial,sans-serif;font-size:17px;text-align:center;vertical-align:middle">${changePhotoLabel}</div>`;

    return `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:728px;border-collapse:collapse;background:#fff;font-family:Arial,sans-serif"><tr><td width="138" style="width:138px;padding:0 18px 0 0;vertical-align:top">${photo}<div style="padding-top:11px;text-align:center"><img src="${escapeHtml(assetUrl(assetBase, school.logoAsset))}" width="124" alt="${escapeHtml(school.label)}" style="display:block;width:124px;height:auto;border:0" /></div><div style="padding-top:12px"><img src="${escapeHtml(assetUrl(assetBase, "social-icons.png"))}" width="118" alt="Social media" style="display:block;width:118px;height:auto;border:0" /></div></td><td style="padding:2px 0 0;vertical-align:top"><div style="color:#666;font-size:25px;font-weight:700;line-height:29px">${escapeHtml(data.name)}</div><div style="margin-top:2px;color:#666;font-size:16px;line-height:20px">${escapeHtml(data.title)}</div><div style="margin-top:2px;color:#666;font-size:16px;line-height:20px">${escapeHtml(data.organization)}</div><div style="height:12px;border-bottom:2px solid #d71920"></div><table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-top:9px;border-collapse:collapse">${infoRow("icon-phone.png", data.phone, data.phone ? `tel:${String(data.phone).replace(/\s/g, "")}` : "", assetBase)}${infoRow("icon-email.png", data.email, data.email ? `mailto:${data.email}` : "", assetBase)}${infoRow("icon-web.png", data.website, normalizedUrl(data.website), assetBase)}${infoRow("icon-address.png", data.address, "", assetBase)}</table></td></tr><tr><td colspan="2" style="padding-top:12px"><img src="${escapeHtml(assetUrl(assetBase, "brand-family.png"))}" width="700" alt="EtonHouse family of schools" style="display:block;width:700px;height:auto;border:0" /></td></tr></table>`;
  }

  function complete(event) {
    try { event.completed(); } catch (_) { /* Outlook may already have closed the item. */ }
  }

  function setSignature(profile, photoFileName, event) {
    Office.context.mailbox.item.body.setSignatureAsync(
      buildSignatureHtml(profile, photoFileName),
      { coercionType: Office.CoercionType.Html },
      function () { complete(event); }
    );
  }

  function addPhotoAndSetSignature(profile, event) {
    const photoFileName = "etonhouse-profile.jpg";
    const base64 = String(profile.photoDataUrl || "").split(",")[1];
    if (!base64) {
      setSignature(profile, "", event);
      return;
    }
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      base64,
      photoFileName,
      { isInline: true },
      function (result) {
        setSignature(profile, result.status === Office.AsyncResultStatus.Succeeded ? photoFileName : "", event);
      }
    );
  }

  function showSetupNotice() {
    const item = Office.context.mailbox.item;
    if (!item || !item.notificationMessages) return;
    item.notificationMessages.replaceAsync("etonhouse-signature-setup", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "请打开 EtonHouse Signature，填写并保存你的邮件签名。",
      icon: "Icon.16x16",
      persistent: false,
    });
  }

  function checkSignature(event) {
    try {
      const profile = Office.context.roamingSettings.get(profileKey);
      if (!profile || !profile.name || !profile.email) {
        showSetupNotice();
        complete(event);
        return;
      }
      addPhotoAndSetSignature(profile, event);
    } catch (_) {
      complete(event);
    }
  }

  if (typeof Office !== "undefined" && Office.actions) {
    Office.actions.associate("checkSignature", checkSignature);
  }
})();
