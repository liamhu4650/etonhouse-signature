(function () {
  const assetRevision = "20260911-new-template";
  const schools = {
    hq: {
      label: "HQ 总部",
      description: "EtonHouse International Education Group",
      logoAsset: "school-hq.png",
      logoWidth: 184,
    },
    ebridge: {
      label: "E-Bridge Pre-School",
      description: "Member of EtonHouse International Education Group",
      logoAsset: "school-ebridge.png",
      logoWidth: 184,
    },
    international: {
      label: "International School",
      description: "EtonHouse International School",
      logoAsset: "school-international.png",
      logoWidth: 196,
    },
    preschool: {
      label: "EtonHouse Pre-School",
      description: "EtonHouse International Pre-School",
      logoAsset: "school-preschool.png",
      logoWidth: 188,
    },
  };

  function escapeHtml(value) {
    return String(value || "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function normalizedUrl(value) {
    const clean = String(value || "").trim();
    if (!clean) return "";
    return /^https?:\/\//i.test(clean) ? clean : `https://${clean}`;
  }

  function assetUrl(base, fileName) {
    return `${String(base || "assets/").replace(/\/?$/, "/")}${fileName}?v=${assetRevision}`;
  }

  function infoRow(iconAsset, value, href, assetBase) {
    if (!value) return "";
    const content = href
      ? `<a href="${escapeHtml(href)}" style="color:#666666;text-decoration:none">${escapeHtml(value)}</a>`
      : escapeHtml(value);
    return `<tr>
      <td width="16" style="width:16px;padding:2px 0 1px 2px;vertical-align:middle"><img src="${escapeHtml(assetUrl(assetBase, iconAsset))}" width="10" height="10" alt="" style="display:block;width:10px;height:10px;border:0" /></td>
      <td style="padding:0;color:#666666;font-family:'Times New Roman',Times,serif;font-size:10pt;font-weight:400;line-height:12pt;vertical-align:middle">${content}</td>
    </tr>`;
  }

  function buildSignatureHtml(data, options = {}) {
    const school = schools[data.school] || schools.hq;
    const assetBase = options.assetBaseUrl || data.assetBaseUrl || "assets/";
    const photoSrc = options.photoCid ? `cid:${options.photoCid}` : data.photoDataUrl;
    const photo = photoSrc
      ? `<img src="${escapeHtml(photoSrc)}" width="128" height="128" alt="${escapeHtml(data.name)}" style="display:block;width:128px;height:128px;object-fit:cover;border:0" />`
      : `<div style="display:block;width:128px;height:128px;background:#156082;font-size:0;line-height:0">&nbsp;</div>`;

    const websiteUrl = normalizedUrl(data.website);
    return `<table role="presentation" width="760" cellpadding="0" cellspacing="0" border="0" style="width:760px;border-collapse:collapse;background:#ffffff;font-family:'Times New Roman',Times,serif">
      <tr>
        <td width="12" height="148" style="width:12px;height:148px;font-size:0;line-height:0">&nbsp;</td>
        <td width="128" height="148" style="width:128px;height:148px;padding:14px 0 6px;vertical-align:top">
          ${photo}
        </td>
        <td width="20" height="148" style="width:20px;height:148px;font-size:0;line-height:0">&nbsp;</td>
        <td width="348" height="148" style="width:348px;height:148px;padding:5px 0 0;vertical-align:top">
          <div style="color:#666666;font-family:'Times New Roman',Times,serif;font-size:16pt;font-weight:700;line-height:19pt;white-space:nowrap">${escapeHtml(data.name)}</div>
          <div style="color:#666666;font-family:'Times New Roman',Times,serif;font-size:12pt;font-weight:400;line-height:15pt;white-space:nowrap">${escapeHtml(data.title)}</div>
          <div style="color:#666666;font-family:'Times New Roman',Times,serif;font-size:12pt;font-weight:400;line-height:15pt;white-space:nowrap">${escapeHtml(data.organization)}</div>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-top:6px;border-collapse:collapse">
            ${infoRow("icon-phone.png", data.phone, data.phone ? `tel:${String(data.phone).replace(/\s/g, "")}` : "", assetBase)}
            ${infoRow("icon-email.png", data.email, data.email ? `mailto:${data.email}` : "", assetBase)}
            ${infoRow("icon-web.png", data.website, websiteUrl, assetBase)}
            ${infoRow("icon-address.png", data.address, "", assetBase)}
          </table>
        </td>
        <td width="252" height="148" style="width:252px;height:148px;padding:8px 0 0;vertical-align:middle">
          <img src="${escapeHtml(assetUrl(assetBase, school.logoAsset))}" width="${school.logoWidth}" alt="${escapeHtml(school.label)}" style="display:block;width:${school.logoWidth}px;height:auto;border:0" />
        </td>
      </tr>
      <tr>
        <td colspan="5" style="padding:0">
          <img src="${escapeHtml(assetUrl(assetBase, "brand-family.png"))}" width="760" height="47" alt="EtonHouse family of schools" style="display:block;width:760px;height:47px;border:0" />
        </td>
      </tr>
    </table>`;
  }

  window.EtonSignature = { schools, buildSignatureHtml };
})();
