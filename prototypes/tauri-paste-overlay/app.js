const root = document.documentElement;
const overlay = document.querySelector(".paste-overlay");
const searchField = document.querySelector("#searchField");
const searchInput = document.querySelector("#searchInput");
const cards = [...document.querySelectorAll(".clip-card")];
const groupTabList = document.querySelector(".group-tab-list");
let tabs = [...document.querySelectorAll(".category-tabs button:not(.new-group-button)")];
const newGroupButton = document.querySelector(".new-group-button");

const controls = {
  width: document.querySelector("#widthControl"),
  opacity: document.querySelector("#opacityControl"),
  blur: document.querySelector("#blurControl"),
  card: document.querySelector("#cardControl"),
};

const outputs = {
  width: document.querySelector("#widthValue"),
  opacity: document.querySelector("#opacityValue"),
  blur: document.querySelector("#blurValue"),
  card: document.querySelector("#cardValue"),
};

const presets = {
  paste: { width: 90, opacity: 60, blur: 10, card: 244 },
  soft: { width: 92, opacity: 76, blur: 24, card: 252 },
  dense: { width: 94, opacity: 88, blur: 12, card: 204 },
};

let selectedIndex = Math.max(
  0,
  cards.findIndex((card) => card.classList.contains("selected")),
);

function openSearch() {
  overlay.classList.add("search-open");
  requestAnimationFrame(() => searchInput.focus());
}

function closeSearch() {
  overlay.classList.remove("search-open");
  searchInput.blur();
}

function setVars() {
  const width = Number(controls.width.value);
  const opacity = Number(controls.opacity.value) / 100;
  const blur = Number(controls.blur.value);
  const card = Number(controls.card.value);

  root.style.setProperty("--panel-width", `${width}vw`);
  root.style.setProperty("--panel-alpha", opacity.toFixed(2));
  root.style.setProperty("--panel-blur", `${blur}px`);
  root.style.setProperty("--card-width", `${card}px`);

  outputs.width.textContent = `${width}%`;
  outputs.opacity.textContent = opacity.toFixed(2);
  outputs.blur.textContent = `${blur}px`;
  outputs.card.textContent = `${card}px`;
}

function selectCard(index) {
  selectedIndex = Math.max(0, Math.min(cards.length - 1, index));
  cards.forEach((card, cardIndex) => {
    card.classList.toggle("selected", cardIndex === selectedIndex);
  });
  cards[selectedIndex].scrollIntoView({
    behavior: "smooth",
    block: "nearest",
    inline: "center",
  });
}

function bindTabs() {
  tabs = [...document.querySelectorAll(".category-tabs button:not(.new-group-button)")];
  tabs.forEach((tab) => {
    if (tab.dataset.bound === "true") {
      return;
    }
    tab.dataset.bound = "true";
    tab.addEventListener("click", () => {
      tabs.forEach((item) => item.classList.remove("active"));
      tab.classList.add("active");
    });
  });
}

function addGroupTab() {
  const groupCount = tabs.length;
  const button = document.createElement("button");
  button.dataset.tab = `group-${groupCount}`;
  button.innerHTML =
    '<span class="group-dot ideas"></span><span class="group-label">&#26032;&#35215;&#12464;&#12523;&#12540;&#12503;</span>';
  groupTabList.append(button);
  bindTabs();
  button.click();
  button.scrollIntoView({
    behavior: "smooth",
    block: "nearest",
    inline: "center",
  });
}

for (const control of Object.values(controls)) {
  control.addEventListener("input", setVars);
}

searchField.addEventListener("click", () => openSearch());
searchInput.addEventListener("focus", () => overlay.classList.add("search-open"));

document.querySelector("#resetButton").addEventListener("click", () => {
  applyPreset("paste");
});

document.querySelectorAll("[data-preset]").forEach((button) => {
  button.addEventListener("click", () => applyPreset(button.dataset.preset));
});

function applyPreset(name) {
  const preset = presets[name] ?? presets.paste;
  controls.width.value = preset.width;
  controls.opacity.value = preset.opacity;
  controls.blur.value = preset.blur;
  controls.card.value = preset.card;
  setVars();
}

cards.forEach((card, index) => {
  card.addEventListener("click", () => selectCard(index));
});

bindTabs();
newGroupButton.addEventListener("click", addGroupTab);

document.addEventListener("keydown", (event) => {
  const isSearchShortcut = (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "f";
  if (isSearchShortcut) {
    event.preventDefault();
    openSearch();
    return;
  }

  if (event.key === "ArrowRight") {
    event.preventDefault();
    selectCard(selectedIndex + 1);
  }
  if (event.key === "ArrowLeft") {
    event.preventDefault();
    selectCard(selectedIndex - 1);
  }
  if (event.key === "Escape") {
    if (overlay.classList.contains("search-open")) {
      closeSearch();
      return;
    }
    overlay.classList.toggle("is-dimmed");
  }
});

setVars();
