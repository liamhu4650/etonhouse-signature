(function () {
  const form = document.getElementById("signature-form");
  const schoolOptions = document.getElementById("school-options");
  const preview = document.getElementById("signature-preview");
  const photoInput = document.getElementById("photo-input");
  const photoThumb = document.getElementById("photo-thumb");
  const photoPlaceholder = document.getElementById("photo-placeholder");
  const removePhoto = document.getElementById("remove-photo");
  const saveButton = document.getElementById("save-button");
  const saveMessage = document.getElementById("save-message");
  const connectionStatus = document.getElementById("connection-status");
  const profileKey = "etonhouse.signature.profile.v1";
  const previewDraftKey = "etonhouse-signature-preview.v2";
  const languageKey = "etonhouse-signature-language";

  const translations = {
    zh: {
      languageAria: "界面语言",
      selectSchoolTitle: "选择所属学校",
      selectSchoolHelp: "模板决定左侧品牌标识，个人资料保持一致。",
      schoolAria: "学校类型",
      detailsTitle: "填写个人资料",
      detailsHelp: "所有资料仅用于生成你的邮件签名。",
      photoAlt: "个人照片预览",
      uploadPhoto: "上传照片",
      squarePhoto: "建议使用正方形证件照",
      personalPhoto: "个人照片",
      photoHelp: "上传后会自动居中裁剪并压缩，确保可以随 M365 签名同步。",
      removePhoto: "移除照片",
      nameLabel: "姓名",
      titleLabel: "职位 / Title",
      organizationLabel: "学校或机构名称",
      phoneLabel: "电话",
      emailLabel: "邮箱",
      websiteLabel: "官方网站",
      addressLabel: "地址",
      previewTitle: "实时预览",
      previewSize: "邮件显示宽度 · 760 px",
      safeTitle: "Outlook 安全排版",
      safeCopy: "签名使用邮件兼容的表格与行内样式生成，不依赖截图。",
      browserStatus: "浏览器预览",
      outlookStatus: "已连接 Microsoft 365",
      saveBrowserTitle: "保存我的签名",
      saveBrowserHelp: "当前保存到浏览器；接入 Outlook 后保存到 M365 邮箱",
      saveOutlookTitle: "保存并应用到 Outlook",
      saveOutlookHelp: "配置将保存在你的 M365 邮箱，并用于所有 Outlook 客户端",
      savedLocal: "已保存浏览器预览。接入 Outlook 后，这里会改为保存至你的 M365 邮箱。",
      savedOutlook: "已保存到你的 M365 邮箱。下一封新邮件将自动使用此签名。",
      saveFailed: "保存失败：",
      schoolLabels: { hq: "HQ 总部", ebridge: "E-Bridge Pre-School", international: "International School", preschool: "EtonHouse Pre-School" },
    },
    en: {
      languageAria: "Interface language",
      selectSchoolTitle: "Choose your school",
      selectSchoolHelp: "The template controls the brand identity while your personal details stay consistent.",
      schoolAria: "School type",
      detailsTitle: "Enter your details",
      detailsHelp: "Your information is used only to create your email signature.",
      photoAlt: "Personal photo preview",
      uploadPhoto: "Upload photo",
      squarePhoto: "A square headshot is recommended",
      personalPhoto: "Personal photo",
      photoHelp: "Your photo is automatically centred, cropped and compressed for M365 signature sync.",
      removePhoto: "Remove photo",
      nameLabel: "Name",
      titleLabel: "Title",
      organizationLabel: "School or organisation",
      phoneLabel: "Phone",
      emailLabel: "Email",
      websiteLabel: "Official website",
      addressLabel: "Address",
      previewTitle: "Live preview",
      previewSize: "Email display width · 760 px",
      safeTitle: "Outlook-safe layout",
      safeCopy: "The signature uses email-compatible tables and inline styles rather than a screenshot.",
      browserStatus: "Browser preview",
      outlookStatus: "Connected to Microsoft 365",
      saveBrowserTitle: "Save my signature",
      saveBrowserHelp: "Saved in this browser for now; in Outlook it will be saved to your M365 mailbox",
      saveOutlookTitle: "Save and apply to Outlook",
      saveOutlookHelp: "Your settings will be saved in your M365 mailbox and used across Outlook clients",
      savedLocal: "Browser preview saved. In Outlook, your signature will be saved to your M365 mailbox.",
      savedOutlook: "Saved to your M365 mailbox. Your next new message will use this signature automatically.",
      saveFailed: "Save failed: ",
      schoolLabels: { hq: "HQ", ebridge: "E-Bridge Pre-School", international: "International School", preschool: "EtonHouse Pre-School" },
    },
  };

  const defaults = {
    school: "hq",
    photoDataUrl: "",
  };
  let state = loadDraft();
  let outlookConnected = false;
  let language = localStorage.getItem(languageKey) || (navigator.language.toLowerCase().startsWith("zh") ? "zh" : "en");
  if (!translations[language]) language = "zh";

  function textFor(key) {
    return translations[language][key];
  }

  function updateConnectionUi() {
    connectionStatus.querySelector("b").textContent = outlookConnected ? textFor("outlookStatus") : textFor("browserStatus");
    saveButton.querySelector("span").textContent = outlookConnected ? textFor("saveOutlookTitle") : textFor("saveBrowserTitle");
    saveButton.querySelector("small").textContent = outlookConnected ? textFor("saveOutlookHelp") : textFor("saveBrowserHelp");
  }

  function applyLanguage(nextLanguage) {
    language = translations[nextLanguage] ? nextLanguage : "zh";
    localStorage.setItem(languageKey, language);
    document.documentElement.lang = language === "zh" ? "zh-CN" : "en";
    document.querySelectorAll("[data-i18n]").forEach((element) => {
      element.textContent = textFor(element.dataset.i18n);
    });
    document.querySelectorAll("[data-i18n-aria]").forEach((element) => {
      element.setAttribute("aria-label", textFor(element.dataset.i18nAria));
    });
    document.querySelectorAll("[data-i18n-alt]").forEach((element) => {
      element.setAttribute("alt", textFor(element.dataset.i18nAlt));
    });
    document.querySelectorAll("[data-language]").forEach((button) => {
      button.setAttribute("aria-pressed", String(button.dataset.language === language));
    });
    updateConnectionUi();
    renderSchoolOptions();
    renderPreview();
  }

  function loadDraft() {
    try {
      return { ...defaults, ...JSON.parse(localStorage.getItem(previewDraftKey) || "{}") };
    } catch (_) {
      return { ...defaults };
    }
  }

  function renderSchoolOptions() {
    schoolOptions.innerHTML = Object.entries(window.EtonSignature.schools)
      .map(([key, school]) => `<button type="button" class="school-card" role="radio" data-school="${key}" aria-checked="${state.school === key}">
        <strong>${textFor("schoolLabels")[key]}</strong><small>${school.description}</small>
      </button>`).join("");
  }

  function applyDataToForm(data) {
    if (!data) return;
    state = { ...state, ...data };
    for (const element of form.elements) {
      if (element.name && Object.prototype.hasOwnProperty.call(data, element.name)) {
        element.value = data[element.name] || "";
      }
    }
    renderSchoolOptions();
    syncPhotoUi();
    if (data.uiLanguage && translations[data.uiLanguage] && data.uiLanguage !== language) {
      applyLanguage(data.uiLanguage);
    } else {
      renderPreview();
    }
  }

  function currentData() {
    return {
      school: state.school,
      photoDataUrl: state.photoDataUrl,
      uiLanguage: language,
      assetBaseUrl: new URL("assets/", window.location.href).href,
      ...Object.fromEntries(new FormData(form).entries()),
    };
  }

  function renderPreview() {
    preview.innerHTML = window.EtonSignature.buildSignatureHtml(currentData());
  }

  async function cropAndCompress(file) {
    const bitmap = await createImageBitmap(file);
    const size = Math.min(bitmap.width, bitmap.height);
    const sourceX = Math.round((bitmap.width - size) / 2);
    const sourceY = Math.round((bitmap.height - size) / 2);
    const canvas = document.createElement("canvas");
    canvas.width = 118;
    canvas.height = 118;
    const context = canvas.getContext("2d", { alpha: false });
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, 118, 118);
    context.drawImage(bitmap, sourceX, sourceY, size, size, 0, 0, 118, 118);
    bitmap.close();

    let quality = 0.78;
    let value = canvas.toDataURL("image/jpeg", quality);
    while (value.length > 18000 && quality > 0.42) {
      quality -= 0.08;
      value = canvas.toDataURL("image/jpeg", quality);
    }
    return value;
  }

  function syncPhotoUi() {
    if (state.photoDataUrl) {
      photoThumb.src = state.photoDataUrl;
      photoThumb.hidden = false;
      photoPlaceholder.hidden = true;
      removePhoto.hidden = false;
    } else {
      photoThumb.removeAttribute("src");
      photoThumb.hidden = true;
      photoPlaceholder.hidden = false;
      removePhoto.hidden = true;
    }
  }

  schoolOptions.addEventListener("click", (event) => {
    const button = event.target.closest("[data-school]");
    if (!button) return;
    state.school = button.dataset.school;
    renderSchoolOptions();
    renderPreview();
  });

  form.addEventListener("input", renderPreview);

  photoInput.addEventListener("change", async () => {
    const file = photoInput.files && photoInput.files[0];
    if (!file) return;
    state.photoDataUrl = await cropAndCompress(file);
    syncPhotoUi();
    renderPreview();
    photoInput.value = "";
  });

  removePhoto.addEventListener("click", () => {
    state.photoDataUrl = "";
    syncPhotoUi();
    renderPreview();
  });

  function saveToOutlook(data) {
    Office.context.roamingSettings.set(profileKey, data);
    Office.context.roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        saveMessage.textContent = textFor("savedOutlook");
      } else {
        saveMessage.textContent = `${textFor("saveFailed")}${result.error.message}`;
      }
    });
  }

  saveButton.addEventListener("click", () => {
    if (!form.reportValidity()) return;
    const data = currentData();
    localStorage.setItem(previewDraftKey, JSON.stringify(data));
    if (outlookConnected) {
      saveToOutlook(data);
    } else {
      saveMessage.textContent = textFor("savedLocal");
    }
    window.setTimeout(() => { saveMessage.textContent = ""; }, 4200);
  });

  document.querySelector(".language-switch").addEventListener("click", (event) => {
    const button = event.target.closest("[data-language]");
    if (button) applyLanguage(button.dataset.language);
  });

  applyDataToForm(state);
  applyLanguage(language);

  if (window.Office && Office.onReady) {
    Office.onReady((info) => {
      if (!info || info.host !== Office.HostType.Outlook) return;
      outlookConnected = true;
      updateConnectionUi();
      const profile = Office.context.roamingSettings.get(profileKey);
      if (profile) {
        applyDataToForm(profile);
      } else {
        const user = Office.context.mailbox.userProfile;
        const emailInput = form.elements.namedItem("email");
        const nameInput = form.elements.namedItem("name");
        if (user && user.emailAddress && emailInput) emailInput.value = user.emailAddress;
        if (user && user.displayName && nameInput) nameInput.value = user.displayName;
        renderPreview();
      }
    });
  }
})();
