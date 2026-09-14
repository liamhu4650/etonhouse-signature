/* Self-contained because classic Outlook loads one event-runtime script. */
(function () {
  const profileKey = "etonhouse.signature.profile.v1";
  const productionAssetFallback = "https://liamhu4650.github.io/etonhouse-signature/assets/";
  const assetRevision = "20260911-new-template";
  const schools = {
    hq: { label: "HQ 总部", logoAsset: "school-hq.png", logoWidth: 184 },
    ebridge: { label: "E-Bridge Pre-School", logoAsset: "school-ebridge.png", logoWidth: 184 },
    international: { label: "International School", logoAsset: "school-international.png", logoWidth: 196 },
    preschool: { label: "EtonHouse Pre-School", logoAsset: "school-preschool.png", logoWidth: 188 },
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
    return `${String(base || productionAssetFallback).replace(/\/?$/, "/")}${fileName}?v=${assetRevision}`;
  }

  function infoRow(iconAsset, value, href, assetBase, alignToFirstLine) {
    if (!value) return "";
    const content = href
      ? `<a href="${escapeHtml(href)}" style="color:#666666;text-decoration:none">${escapeHtml(value)}</a>`
      : escapeHtml(value);
    const iconPadding = alignToFirstLine ? "3px 0 1px 2px" : "2px 0 1px 2px";
    const verticalAlign = alignToFirstLine ? "top" : "middle";
    return `<tr><td width="16" style="width:16px;padding:${iconPadding};vertical-align:${verticalAlign}"><img src="${escapeHtml(assetUrl(assetBase, iconAsset))}" width="10" height="10" alt="" style="display:block;width:10px;height:10px;border:0" /></td><td style="padding:0;color:#666666;font-family:'Times New Roman',Times,serif;font-size:10pt;font-weight:400;line-height:12pt;vertical-align:${verticalAlign}">${content}</td></tr>`;
  }

  function buildSignatureHtml(data, photoFileName) {
    const school = schools[data.school] || schools.hq;
    const assetBase = data.assetBaseUrl || productionAssetFallback;
    const photo = photoFileName
      ? `<img src="cid:${escapeHtml(photoFileName)}" width="128" height="128" alt="${escapeHtml(data.name)}" style="display:block;width:128px;height:128px;border:0" />`
      : `<div style="display:block;width:128px;height:128px;background:#156082;font-size:0;line-height:0">&nbsp;</div>`;

    return `<table role="presentation" width="760" cellpadding="0" cellspacing="0" border="0" style="width:760px;border-collapse:collapse;background:#ffffff;font-family:'Times New Roman',Times,serif"><tr><td width="12" height="148" style="width:12px;height:148px;font-size:0;line-height:0">&nbsp;</td><td width="128" height="148" style="width:128px;height:148px;padding:14px 0 6px;vertical-align:top">${photo}</td><td width="20" height="148" style="width:20px;height:148px;font-size:0;line-height:0">&nbsp;</td><td width="348" height="148" style="width:348px;height:148px;padding:5px 0 0;vertical-align:top"><div style="color:#666666;font-family:'Times New Roman',Times,serif;font-size:16pt;font-weight:700;line-height:19pt;white-space:nowrap">${escapeHtml(data.name)}</div><div style="color:#666666;font-family:'Times New Roman',Times,serif;font-size:12pt;font-weight:400;line-height:15pt;white-space:nowrap">${escapeHtml(data.title)}</div><div style="color:#666666;font-family:'Times New Roman',Times,serif;font-size:12pt;font-weight:400;line-height:15pt;white-space:nowrap">${escapeHtml(data.organization)}</div><table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-top:6px;border-collapse:collapse">${infoRow("icon-phone.png", data.phone, data.phone ? `tel:${String(data.phone).replace(/\s/g, "")}` : "", assetBase)}${infoRow("icon-email.png", data.email, data.email ? `mailto:${data.email}` : "", assetBase)}${infoRow("icon-web.png", data.website, normalizedUrl(data.website), assetBase)}${infoRow("icon-address.png", data.address, "", assetBase, true)}</table></td><td width="252" height="148" style="width:252px;height:148px;padding:8px 0 0;vertical-align:middle"><img src="${escapeHtml(assetUrl(assetBase, school.logoAsset))}" width="${school.logoWidth}" alt="${escapeHtml(school.label)}" style="display:block;width:${school.logoWidth}px;height:auto;border:0" /></td></tr><tr><td colspan="5" style="padding:0"><img src="${escapeHtml(assetUrl(assetBase, "brand-family.png"))}" width="760" height="47" alt="EtonHouse family of schools" style="display:block;width:760px;height:47px;border:0" /></td></tr></table>`;
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
