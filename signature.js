(function () {
  const schools = {
    hq: {
      label: "HQ 总部",
      description: "EtonHouse International Education Group",
      brandTop: "EtonHouse",
      brandBottom: "International Education Group",
      logoAsset: "school-hq.png",
    },
    ebridge: {
      label: "E-Bridge Pre-School",
      description: "Member of EtonHouse International Education Group",
      brandTop: "E-Bridge Pre-School",
      brandBottom: "Member of EtonHouse International Education Group",
      logoAsset: "school-ebridge.png",
    },
    international: {
      label: "International School",
      description: "EtonHouse International School",
      brandTop: "EtonHouse",
      brandBottom: "International School",
      logoAsset: "school-international.png",
    },
    preschool: {
      label: "EtonHouse Pre-School",
      description: "EtonHouse International Pre-School",
      brandTop: "EtonHouse",
      brandBottom: "International Pre-School",
      logoAsset: "school-preschool.png",
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
    return `${String(base || "assets/").replace(/\/?$/, "/")}${fileName}`;
  }

  function infoRow(iconAsset, value, href, assetBase) {
    if (!value) return "";
    const content = href
      ? `<a href="${escapeHtml(href)}" style="color:#6b6b6b;text-decoration:none">${escapeHtml(value)}</a>`
      : escapeHtml(value);
    return `<tr>
      <td width="27" style="width:27px;padding:2px 0;vertical-align:top"><img src="${escapeHtml(assetUrl(assetBase, iconAsset))}" width="20" alt="" style="display:block;width:20px;height:auto;border:0" /></td>
      <td style="padding:2px 0;color:#6b6b6b;font-family:Arial,sans-serif;font-size:12px;font-weight:700;line-height:20px;vertical-align:top">${content}</td>
    </tr>`;
  }

  function buildSignatureHtml(data, options = {}) {
    const school = schools[data.school] || schools.hq;
    const assetBase = options.assetBaseUrl || data.assetBaseUrl || "assets/";
    const photoSrc = options.photoCid ? `cid:${options.photoCid}` : data.photoDataUrl;
    const changePhotoLabel = data.uiLanguage === "en" ? "CHANGE PHOTO" : "更换照片";
    const photo = photoSrc
      ? `<img src="${escapeHtml(photoSrc)}" width="118" height="118" alt="${escapeHtml(data.name)}" style="display:block;width:118px;height:118px;object-fit:cover;border:0" />`
      : `<div style="display:table-cell;width:118px;height:118px;color:#fff;background:#d71920;font-family:Microsoft YaHei,Arial,sans-serif;font-size:17px;text-align:center;vertical-align:middle">${changePhotoLabel}</div>`;

    const websiteUrl = normalizedUrl(data.website);
    return `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:728px;border-collapse:collapse;background:#fff;font-family:Arial,sans-serif">
      <tr>
        <td width="138" style="width:138px;padding:0 18px 0 0;vertical-align:top">
          ${photo}
          <div style="padding-top:11px;text-align:center"><img src="${escapeHtml(assetUrl(assetBase, school.logoAsset))}" width="124" alt="${escapeHtml(school.label)}" style="display:block;width:124px;height:auto;border:0" /></div>
          <div style="padding-top:12px"><img src="${escapeHtml(assetUrl(assetBase, "social-icons.png"))}" width="118" alt="Social media" style="display:block;width:118px;height:auto;border:0" /></div>
        </td>
        <td style="padding:2px 0 0;vertical-align:top">
          <div style="color:#666;font-size:25px;font-weight:700;line-height:29px">${escapeHtml(data.name)}</div>
          <div style="margin-top:2px;color:#666;font-size:16px;line-height:20px">${escapeHtml(data.title)}</div>
          <div style="margin-top:2px;color:#666;font-size:16px;line-height:20px">${escapeHtml(data.organization)}</div>
          <div style="height:12px;border-bottom:2px solid #d71920"></div>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-top:9px;border-collapse:collapse">
            ${infoRow("icon-phone.png", data.phone, data.phone ? `tel:${String(data.phone).replace(/\s/g, "")}` : "", assetBase)}
            ${infoRow("icon-email.png", data.email, data.email ? `mailto:${data.email}` : "", assetBase)}
            ${infoRow("icon-web.png", data.website, websiteUrl, assetBase)}
            ${infoRow("icon-address.png", data.address, "", assetBase)}
          </table>
        </td>
      </tr>
      <tr>
        <td colspan="2" style="padding-top:12px">
          <img src="${escapeHtml(assetUrl(assetBase, "brand-family.png"))}" width="700" alt="EtonHouse family of schools" style="display:block;width:700px;height:auto;border:0" />
        </td>
      </tr>
    </table>`;
  }

  window.EtonSignature = { schools, buildSignatureHtml };
})();
